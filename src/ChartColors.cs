using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace MelloCharts {
    public sealed class ChartColors {
        public string actual,forecast,baseline,increase,decrease,background;
        public List<string> categories;
    }
    public static class ColorTools {
        static readonly string[] Fields={"actual","forecast","baseline","increase","decrease","background"};
        public static string Normalize(string input) {
            if(input==null || input.Length>60)throw new ArgumentException("Informe HEX (#RRGGBB) ou RGB (0 a 255).");
            string text=input.Trim();
            if(Regex.IsMatch(text,@"^#?[0-9a-fA-F]{6}$"))return "#"+text.TrimStart('#').ToUpperInvariant();
            if(Regex.IsMatch(text,@"^#[0-9a-fA-F]{3}$"))return "#"+String.Concat(text.Substring(1).Select(c=>new string(c,2))).ToUpperInvariant();
            Match rgb=Regex.Match(text,@"^(?:rgb\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*\)|([0-9]{1,3})\s*[,;]\s*([0-9]{1,3})\s*[,;]\s*([0-9]{1,3}))$",RegexOptions.IgnoreCase);
            if(rgb.Success) {
                int start=rgb.Groups[1].Success?1:4;
                int[] values=Enumerable.Range(start,3).Select(i=>Int32.Parse(rgb.Groups[i].Value,CultureInfo.InvariantCulture)).ToArray();
                if(values.All(v=>v<=255))return "#"+String.Concat(values.Select(v=>v.ToString("X2",CultureInfo.InvariantCulture)));
            }
            throw new ArgumentException("Cor inválida. Use #RRGGBB, #RGB ou RGB com três valores de 0 a 255.");
        }
        public static ChartColors Clone(ChartColors colors) {
            if(colors==null)return null;
            return new ChartColors {actual=colors.actual,forecast=colors.forecast,baseline=colors.baseline,increase=colors.increase,decrease=colors.decrease,background=colors.background,categories=colors.categories==null?null:colors.categories.ToList()};
        }
        public static void Validate(ChartColors colors) {
            if(colors==null)return;
            foreach(string key in Fields) {
                string value=(string)typeof(ChartColors).GetField(key).GetValue(colors);
                if(value!=null && !Regex.IsMatch(value,@"^#[0-9a-fA-F]{6}$"))throw new ArgumentException("Cor de "+key+" inválida. O projeto usa #RRGGBB.");
            }
            if(colors.categories!=null && (colors.categories.Count<1||colors.categories.Count>12||colors.categories.Any(c=>c==null||!Regex.IsMatch(c,@"^#[0-9a-fA-F]{6}$"))))
                throw new ArgumentException("Use de 1 a 12 cores de categoria no formato #RRGGBB.");
        }
        public static ChartColors Read(object raw) {
            if(raw==null)return null;
            IDictionary<string,object> input=raw as IDictionary<string,object>;
            if(input==null || input.Keys.Any(k=>!Fields.Contains(k)&&k!="categories"))throw new ArgumentException("Objeto de cores personalizadas inválido.");
            ChartColors result=new ChartColors();
            foreach(string key in Fields)if(input.ContainsKey(key)&&input[key]!=null) {
                if(!(input[key] is string))throw new ArgumentException("Cor de "+key+" deve ser um texto #RRGGBB.");
                typeof(ChartColors).GetField(key).SetValue(result,(string)input[key]);
            }
            if(input.ContainsKey("categories")&&input["categories"]!=null) {
                IList list=input["categories"] as IList;
                if(list==null || list.Count>12 || list.Cast<object>().Any(c=>!(c is string)))throw new ArgumentException("Lista de cores de categoria inválida.");
                result.categories=list.Cast<string>().ToList();
            }
            Validate(result);return result;
        }
        static double Luminance(string color) {
            int value=Int32.Parse(color.Substring(1),NumberStyles.HexNumber,CultureInfo.InvariantCulture);
            Func<int,double> linear=c=>{double n=c/255.0;return n<=.04045?n/12.92:Math.Pow((n+.055)/1.055,2.4);};
            return .2126*linear((value>>16)&255)+.7152*linear((value>>8)&255)+.0722*linear(value&255);
        }
        public static double Contrast(string a,string b) {double x=Luminance(a),y=Luminance(b);return (Math.Max(x,y)+.05)/(Math.Min(x,y)+.05);}
        public static string Ink(string fill) {return Contrast("#FFFFFF",fill)>Contrast("#000000",fill)?"#FFFFFF":"#000000";}
        public static string Readable(string requested,string fill) {return Contrast(requested,fill)>=4.5?requested:Ink(fill);}
    }
    public static partial class ChartCore {
        sealed class SurfaceColors {
            public string Background,Ink,Muted,Grid,Axis,Band;
            public SurfaceColors(ChartDocument d) {
                bool dark=d.style=="dark";
                Background=dark?"#101E2E":"#FFFFFF";Ink=dark?"#F3F6FB":"#20374D";Muted=dark?"#A4B6C9":"#65798B";
                Grid=dark?"#26384B":"#E9EFF3";Axis=dark?"#526A82":"#AABAC7";Band=dark?"#192D42":"#F1F6F9";
                if(d.customColors!=null&&d.customColors.background!=null) {
                    Background=d.customColors.background;Ink=ColorTools.Ink(Background);Muted=Ink;
                    Grid=Mix(Background,Ink,.12);Axis=Mix(Background,Ink,.42);Band=Mix(Background,Ink,.05);
                    if(ColorTools.Contrast(Ink,Band)<4.5)Band=Background;
                }
            }
        }
        static Palette Colors(ChartDocument d) {
            Palette p=Colors(d.palette,d.style=="dark");ChartColors c=d.customColors;
            if(c!=null){p.Actual=c.actual??p.Actual;p.Forecast=c.forecast??p.Forecast;p.Baseline=c.baseline??p.Baseline;p.Positive=c.increase??p.Positive;p.Negative=c.decrease??p.Negative;}
            return p;
        }
        static string FillInk(ChartDocument d,string fill,string legacy) {return d.customColors==null?legacy:ColorTools.Ink(fill);}
        static string ColorText(ChartDocument d,string color,string background) {return d.customColors==null?color:ColorTools.Readable(color,background);}
        static string DefaultForecastLine(ChartDocument d) {return d.style=="dark"?"#79C9BA":d.palette=="forest"?"#4C8D75":d.palette=="ink"?"#758DA5":"#478F84";}
        public static ChartColors ColorDefaults(ChartDocument d) {
            Validate(d);Palette p=Colors(d.palette,d.style=="dark");
            return new ChartColors {actual=p.Actual,forecast=(d.type=="lines"||d.type=="area"||d.type=="combo")?DefaultForecastLine(d):p.Forecast,baseline=p.Baseline,increase=p.Positive,decrease=p.Negative,background=d.style=="dark"?"#101E2E":"#FFFFFF",categories=BuiltinSliceColors(d).ToList()};
        }
        public static ChartColors EffectiveColors(ChartDocument d) {
            ChartColors result=ColorDefaults(d),custom=d.customColors;if(custom==null)return result;
            foreach(string key in new[]{"actual","forecast","baseline","increase","decrease","background"}) {
                string value=(string)typeof(ChartColors).GetField(key).GetValue(custom);
                if(value!=null)typeof(ChartColors).GetField(key).SetValue(result,value);
            }
            if(custom.categories!=null)for(int i=0;i<custom.categories.Count;i++)result.categories[i]=custom.categories[i];
            return result;
        }
    }
}
