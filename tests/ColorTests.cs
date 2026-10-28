using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using MelloCharts;

public static class ColorTests {
    static int passed,failed;
    static readonly BindingFlags Private=BindingFlags.Instance|BindingFlags.NonPublic;
    static void Is(bool value,string message){if(!value)throw new Exception(message);}
    static void Test(string name,Action action){try{action();passed++;Console.WriteLine("PASS "+name);}catch(Exception ex){failed++;Console.WriteLine("FAIL "+name+": "+(ex.InnerException??ex).Message);}}
    static void Reject(Action action){bool rejected=false;try{action();}catch(ArgumentException){rejected=true;}catch(TargetInvocationException ex){rejected=ex.InnerException is ArgumentException;}Is(rejected,"Invalid color was accepted");}
    static ChartDocument Custom(string type,string colors){JavaScriptSerializer serializer=new JavaScriptSerializer();var root=(IDictionary<string,object>)serializer.DeserializeObject(ChartCore.Serialize(ChartCore.Example(type)));root["customColors"]=serializer.DeserializeObject(colors);return ChartCore.Deserialize(serializer.Serialize(root));}
    static IDictionary<string,object> SerializedColors(ChartDocument d){var o=new JavaScriptSerializer().DeserializeObject(ChartCore.Serialize(d)) as IDictionary<string,object>;Is(o.ContainsKey("customColors")&&o["customColors"] is IDictionary<string,object>,"Custom colors were discarded");return (IDictionary<string,object>)o["customColors"];}
    static Type ColorTools(){Type t=typeof(ChartCore).Assembly.GetType("MelloCharts.ColorTools");Is(t!=null,"Color input support is missing");return t;}
    static string Normalize(string value){return (string)ColorTools().GetMethod("Normalize").Invoke(null,new object[]{value});}
    static object Field(object owner,string name){return owner.GetType().GetField(name,Private).GetValue(owner);}
    static object Call(object owner,string name,params object[] args){return owner.GetType().GetMethod(name,Private).Invoke(owner,args);}
    static Form ColorForm(ChartDocument d){Type t=typeof(ChartCore).Assembly.GetType("MelloCharts.ChartColorForm");Is(t!=null,"Custom color editor is missing");return (Form)Activator.CreateInstance(t,new object[]{d});}
    static double Luminance(string hex){Color c=ColorTranslator.FromHtml(hex);Func<byte,double> linear=v=>{double n=v/255.0;return n<=.04045?n/12.92:Math.Pow((n+.055)/1.055,2.4);};return .2126*linear(c.R)+.7152*linear(c.G)+.0722*linear(c.B);}
    static double Contrast(string a,string b){double x=Luminance(a),y=Luminance(b);return (Math.Max(x,y)+.05)/(Math.Min(x,y)+.05);}
    [STAThread] public static int Main(){
        Test("HEX and RGB input normalize to exact RGB colors",delegate{
            string[,] cases={{"#0fa","#00FFAA"},{"123abc","#123ABC"},{"#abcdef","#ABCDEF"},{"rgb(18, 52, 86)","#123456"},{"0; 128; 255","#0080FF"},{"255,255,255","#FFFFFF"}};
            for(int i=0;i<cases.GetLength(0);i++)Is(Normalize(cases[i,0])==cases[i,1],"Wrong normalized RGB");
        });
        Test("color input rejects alpha overflow junk and blank values",delegate{
            foreach(string input in new[]{"","#12345","#12345678","rgb(256,0,0)","-1,0,0","#GG0000","rgb(1,2,3)script","1,2,3,4"})Reject(()=>Normalize(input));
        });
        Test("custom series and slice colors survive JSON",delegate{
            var c=SerializedColors(Custom("donut","{\"actual\":\"#FFCC00\",\"forecast\":\"#112233\",\"background\":\"#334455\",\"categories\":[\"#FF0000\",\"#00FF00\"]}"));
            Is((string)c["actual"]=="#FFCC00"&&(string)c["forecast"]=="#112233"&&(string)c["background"]=="#334455","Series color lost");
            Is((string)((IList)c["categories"])[1]=="#00FF00","Slice color lost");
        });
        Test("malformed custom color JSON is rejected",delegate{
            foreach(string json in new[]{"true","[]","{\"actual\":5}","{\"actual\":\"red\"}","{\"actual\":\"#12345678\"}","{\"wrong\":\"#123456\"}","{\"categories\":[]}","{\"categories\":[null]}","{\"categories\":\"#123456\"}","{\"categories\":["+String.Join(",",Enumerable.Repeat("\"#123456\"",13))+"]}"})
                Reject(()=>Custom("pie",json));
        });
        Test("cloned custom slice colors are independent",delegate{
            ChartDocument original=Custom("pie","{\"categories\":[\"#123456\"]}"),copy=ChartCore.Clone(original);
            object settings=typeof(ChartDocument).GetField("customColors").GetValue(copy);
            ((IList)settings.GetType().GetField("categories").GetValue(settings))[0]="#FF0000";
            Is((string)((IList)SerializedColors(original)["categories"])[0]=="#123456","Clone changed original colors");
        });
        Test("old projects retain the default palette",delegate{
            ChartDocument d=ChartCore.Example("buildup");Scene s=ChartCore.BuildScene(ChartCore.Deserialize(ChartCore.Serialize(d)));
            Is(s.Elements.First(e=>e.Role=="bar").Color=="#244F73"&&s.Elements.First(e=>e.Role=="background").Color=="#FFFFFF","Legacy defaults changed");
        });
        Test("custom colors reach every chart type",delegate{
            foreach(var definition in ChartCatalog.All){
                Scene s=ChartCore.BuildScene(Custom(definition.Key,"{\"actual\":\"#9C27B0\",\"forecast\":\"#FFCC00\",\"baseline\":\"#009688\",\"categories\":[\"#E91E63\"]}"));
                bool professional=definition.Key=="gantt"||definition.Key=="scatter"||definition.Key=="bubble";
                string role=definition.Key=="gantt"?"ganttProgress":definition.Key=="scatter"?"point":definition.Key=="bubble"?"pointLabel":definition.Key=="pie"||definition.Key=="donut"?"slice":definition.Key=="lines"||definition.Key=="area"?"actualLine":"bar";
                if(definition.Key=="bubble")Is(s.Elements.Any(e=>e.Role=="bubble"&&e.Color!="#9C27B0"),"Bubble color was not composed with surface");
                else Is(s.Elements.Any(e=>e.Role==role&&e.Color==(role=="slice"||professional?"#E91E63":"#9C27B0")),"Color missing in "+definition.Key);
            }
        });
        Test("waterfall honors custom positive and negative colors",delegate{
            Scene s=ChartCore.BuildScene(Custom("waterfall","{\"increase\":\"#FF8800\",\"decrease\":\"#6622CC\"}"));
            Is(s.Elements.Any(e=>e.Role=="bar"&&e.Category==1&&e.Color=="#FF8800"),"Increase color lost");
            Is(s.Elements.Any(e=>e.Role=="bar"&&e.Category==2&&e.Color=="#6622CC"),"Decrease color lost");
        });
        Test("circular legends and slices share their exact custom colors",delegate{
            foreach(string type in new[]{"pie","donut"}){
                Scene s=ChartCore.BuildScene(Custom(type,"{\"categories\":[\"#012345\",\"#FEDCBA\"]}"));
                foreach(var slice in s.Elements.Where(e=>e.Role=="slice"))Is(s.Elements.Any(e=>e.Role=="legendSwatch"&&e.Category==slice.Category&&e.Color==slice.Color),"Swatch differs from slice");
                Is(s.Elements.First(e=>e.Role=="slice").Color=="#012345","First slice ignored");
            }
        });
        Test("custom projected line matches the corresponding legend",delegate{
            foreach(string type in new[]{"lines","area","combo"}){
                Scene s=ChartCore.BuildScene(Custom(type,"{\"forecast\":\"#AB12EF\"}"));
                Is(s.Elements.First(e=>e.Role=="forecastLine").Color=="#AB12EF","Projected line ignored");
                Is(s.Elements.Where(e=>e.Role=="legend"&&e.Type=="line").Skip(1).First().Color=="#AB12EF","Legend differs");
            }
        });
        Test("custom backgrounds retain readable heading and captions",delegate{
            foreach(var definition in ChartCatalog.All)foreach(string background in new[]{"#FFFFFF","#000000","#808080","#FFFF00","#663399"}){
                Scene s=ChartCore.BuildScene(Custom(definition.Key,"{\"background\":\""+background+"\",\"actual\":\"#FFFF00\"}"));
                Is(s.Elements.First(e=>e.Role=="background").Color==background,"Background ignored "+definition.Key);
                foreach(var e in s.Elements.Where(e=>e.Role=="title"||e.Role=="subtitle"||e.Role=="category"||e.Role=="source"||e.Role=="eyebrow"))
                    Is(Contrast(e.Color,background)>=4.5,"Unreadable "+e.Role+" in "+definition.Key);
            }
        });
        Test("values inside bright and dark fills use readable text",delegate{
            foreach(string type in new[]{"buildup","stacked100"})foreach(string fill in new[]{"#FFFFFF","#FFFF00","#000000","#808080"}){
                Scene s=ChartCore.BuildScene(Custom(type,"{\"actual\":\""+fill+"\",\"forecast\":\""+fill+"\"}"));
                foreach(var e in s.Elements.Where(e=>e.Role=="actualValue"||e.Role=="forecastValue"||e.Role=="percentValue"))
                    Is(Contrast(e.Color,fill)>=4.5,"Unreadable value on "+fill);
            }
        });
        Test("editor and data import preserve custom colors",delegate{
            using(EditorForm editor=new EditorForm(null,Custom("lines","{\"actual\":\"#123456\",\"forecast\":\"#FF9900\"}"),null)){
                ChartDocument next=ChartCore.ImportDelimited(ChartCore.ToCsv(ChartCore.Example("lines")),"lines");Call(editor,"ImportData",next,"Teste");
                var c=SerializedColors((ChartDocument)Call(editor,"ReadEditor"));Is((string)c["forecast"]=="#FF9900","Import discarded custom colors");
            }
        });
        Test("custom color form accepts typed RGB and short HEX",delegate{
            using(Form form=ColorForm(ChartCore.Example("buildup"))){
                ((CheckBox)Field(form,"enabled")).Checked=true;
                var inputs=(Dictionary<string,TextBox>)Field(form,"inputs");inputs["actual"].Text="#f0a";inputs["forecast"].Text="rgb(1, 2, 3)";
                Is((bool)Call(form,"TryAccept"),"Valid RGB rejected");
                object result=form.GetType().GetProperty("Result").GetValue(form,null);
                Is((string)result.GetType().GetField("actual").GetValue(result)=="#FF00AA"&&(string)result.GetType().GetField("forecast").GetValue(result)=="#010203","Input color lost");
            }
        });
        Test("invalid color leaves the color form open with an error",delegate{
            using(Form form=ColorForm(ChartCore.Example("buildup"))){
                ((CheckBox)Field(form,"enabled")).Checked=true;((Dictionary<string,TextBox>)Field(form,"inputs"))["actual"].Text="300,0,0";
                Is(!(bool)Call(form,"TryAccept")&&form.DialogResult!=DialogResult.OK,"Invalid input applied");
                Is(((Label)Field(form,"error")).Text.Length>0,"Missing validation feedback");
            }
        });
        Test("reset removes color overrides without changing data",delegate{
            using(Form form=ColorForm(Custom("buildup","{\"actual\":\"#FF00AA\"}"))){
                Call(form,"ResetDefaults");Is((bool)Call(form,"TryAccept"),"Reset rejected");
                Is(form.GetType().GetProperty("Result").GetValue(form,null)==null,"Reset retained an override");
            }
        });
        Test("PNG export renders the selected background and bar fill",delegate{
            Scene s=ChartCore.BuildScene(Custom("columns","{\"actual\":\"#CC00FF\",\"background\":\"#FFEEDD\"}"));
            using(Bitmap png=ChartPreview.RenderImage(s,1280,720)){
                Is(png.GetPixel(2,2).ToArgb()==ColorTranslator.FromHtml("#FFEEDD").ToArgb(),"PNG background differs");
                var bar=s.Elements.First(e=>e.Role=="bar");Is(png.GetPixel((int)bar.X+4,(int)bar.Y+4).ToArgb()==ColorTranslator.FromHtml("#CC00FF").ToArgb(),"PNG bar differs");
            }
        });
        Test("PowerPoint metadata preserves complete custom colors",delegate{
            ChartDocument d=Custom("donut","{\"actual\":\"#123456\",\"background\":\"#001122\",\"categories\":[\"#ABCDEF\"]}");
            var c=SerializedColors(Metadata.Decode(Metadata.Encode(d)));Is((string)c["background"]=="#001122"&&(string)((IList)c["categories"])[0]=="#ABCDEF","Metadata lost colors");
        });
        Test("highlight bands keep text readable on midtone backgrounds",delegate{
            foreach(string type in new[]{"bars","buildup","pie","donut"}){
                Scene s=ChartCore.BuildScene(Custom(type,"{\"background\":\"#757575\"}"));
                foreach(var band in s.Elements.Where(e=>e.Role=="rowBand"||e.Role=="totalBackground"||e.Role=="legendBand"))
                    foreach(var text in s.Elements.Where(e=>e.Role=="category")){
                        double cx=text.X+text.W/2,cy=text.Y+text.H/2;
                        if(cx>=band.X&&cx<=band.X+band.W&&cy>=band.Y&&cy<=band.Y+band.H)
                            Is(Contrast(text.Color,band.Color)>=4.5,"Highlight reduced text contrast in "+type);
                    }
            }
        });
        Test("Unicode digits are rejected as invalid RGB input",delegate{
            Reject(()=>Normalize("rgb(\u0661,2,3)"));
        });
        Console.WriteLine(passed+" color checks passed; "+failed+" failed");return failed==0?0:1;
    }
}
