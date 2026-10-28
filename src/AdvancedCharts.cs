using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace MelloCharts {
    public static partial class ChartCore {
        sealed class ExecutiveTheme {
            public readonly Palette Palette;
            public readonly string Background, Ink, Muted, Grid, Axis, Band, ForecastLine;
            public ExecutiveTheme(ChartDocument d) {
                Palette=Colors(d);SurfaceColors surface=new SurfaceColors(d);
                Background=surface.Background;Ink=surface.Ink;Muted=surface.Muted;Grid=surface.Grid;Axis=surface.Axis;Band=surface.Band;
                ForecastLine=d.customColors!=null&&d.customColors.forecast!=null?d.customColors.forecast:DefaultForecastLine(d);
            }
        }
        sealed class ExecutiveScale {
            public double Magnitude, Low, High, Step;
            public double Map(double value,double first,double last) {
                return first+(value/Magnitude-Low)/(High-Low)*(last-first);
            }
        }
        static ExecutiveScale Scale(IEnumerable<double> input) {
            List<double> values=input.Concat(new[]{0.0}).ToList();
            double magnitude=values.Max(v=>Math.Abs(v)); if(magnitude==0) magnitude=1;
            double low=values.Min()/magnitude,high=values.Max()/magnitude;
            if(low==high) high=low+1;
            double span=high-low;
            if(low<0) low-=span*.17;
            if(high>0) high+=span*.22;
            double initial=(high-low)/5, power=Math.Pow(10,Math.Floor(Math.Log10(initial))),step=10*power;
            foreach(double n in new[]{1.0,2,2.5,5,10}) if(n*power>=initial) {step=n*power;break;}
            double raw=initial*magnitude,absolutePower=Math.Pow(10,Math.Floor(Math.Log10(raw)));
            if(raw>0 && absolutePower>0) foreach(double n in new[]{1.0,2,2.5,5,10}) if(n*absolutePower>=raw) {
                double candidate=n*absolutePower/magnitude; if(candidate>0 && !Double.IsInfinity(candidate))step=candidate;break;
            }
            return new ExecutiveScale {Magnitude=magnitude,Low=Math.Floor(low/step)*step,High=Math.Ceiling(high/step)*step,Step=step};
        }
        static string ExecutiveValue(ChartDocument d,double value,bool axis) {
            if(value!=0 && Math.Abs(value)<.0001) return value.ToString("0.##E+0",Pt);
            if(axis && Math.Abs(value-Math.Round(value))<=Math.Max(1,Math.Abs(value))*1e-10)value=Math.Round(value);
            int decimals=d.decimals;
            if(axis && value!=Math.Truncate(value)) decimals=Math.Max(decimals,Math.Min(6,(int)Math.Max(1,-Math.Floor(Math.Log10(Math.Abs(value)))+1)));
            return ValueLabel(value,decimals,d.compactNumbers || (axis && Math.Abs(value)>=1000000));
        }
        static string Mix(string background,string foreground,double strength) {
            int a=Int32.Parse(background.Substring(1),NumberStyles.HexNumber,CultureInfo.InvariantCulture);
            int b=Int32.Parse(foreground.Substring(1),NumberStyles.HexNumber,CultureInfo.InvariantCulture);
            int r=(int)Math.Round(((a>>16)&255)*(1-strength)+((b>>16)&255)*strength);
            int g=(int)Math.Round(((a>>8)&255)*(1-strength)+((b>>8)&255)*strength);
            int blue=(int)Math.Round((a&255)*(1-strength)+(b&255)*strength);
            return "#"+((r<<16)|(g<<8)|blue).ToString("X6",CultureInfo.InvariantCulture);
        }
        static void AddEllipse(Scene scene,double x,double y,double width,double height,string color,string role,int category) {
            if(width>0 && height>0)scene.Elements.Add(new SceneElement {Type="ellipse",X=x,Y=y,W=width,H=height,Color=color,Role=role,Category=category});
        }
        static void AddPolygon(Scene scene,IEnumerable<double> coordinates,string color,string role,int category) {
            List<double> p=coordinates.ToList();
            if(p.Count<6 || p.Count%2!=0)throw new InvalidOperationException("Polígono inválido.");
            if(p.Any(v=>Double.IsInfinity(v)||Double.IsNaN(v)))throw new InvalidOperationException("Coordenada inválida.");
            if(p[0]!=p[p.Count-2] || p[1]!=p[p.Count-1]) {p.Add(p[0]);p.Add(p[1]);}
            double minX=Double.MaxValue,maxX=Double.MinValue,minY=Double.MaxValue,maxY=Double.MinValue;
            for(int i=0;i<p.Count;i+=2) {minX=Math.Min(minX,p[i]);maxX=Math.Max(maxX,p[i]);minY=Math.Min(minY,p[i+1]);maxY=Math.Max(maxY,p[i+1]);}
            if(maxX-minX<=0 || maxY-minY<=0)return;
            // Invisible wedges can collapse when Office converts coordinates to single precision.
            // Keep the underlying values and legend, but omit an imperceptible native shape.
            double twiceArea=0;
            for(int i=2;i<p.Count-2;i+=2)
                twiceArea+=(p[i]-p[0])*(p[i+3]-p[1])-(p[i+2]-p[0])*(p[i+1]-p[1]);
            if(Math.Abs(twiceArea)<.02)return;
            scene.Elements.Add(new SceneElement {Type="polygon",Points=p.ToArray(),X=minX,Y=minY,W=maxX-minX,H=maxY-minY,Color=color,Role=role,Category=category});
        }
        static Scene ExecutiveFrame(ChartDocument d,ExecutiveTheme t) {
            Scene s=new Scene();
            AddRect(s,0,0,1280,720,t.Background,"background",-1);
            if(d.style!="minimal") {
                AddRect(s,64,29,28,4,t.Palette.Actual,"accent",-1);
                AddText(s,104,21,1000,24,ChartCatalog.Get(d.type).Name.ToUpperInvariant(),13,ColorText(d,t.Palette.Actual,t.Background),true,"left","eyebrow",-1);
            }
            AddText(s,64,54,1152,78,Wrap(d.title,59),36,t.Ink,true,"left","title",-1);
            AddText(s,64,137,1152,34,Wrap(d.subtitle,100),18,t.Muted,false,"left","subtitle",-1);
            string unit=d.type=="stacked100"?"Participação (%) · base em "+d.unit:d.unit+(d.compactNumbers?" (abreviado)":"");
            AddText(s,64,193,470,26,unit,15,t.Muted,false,"left","unit",-1);
            AddLine(s,64,664,1216,664,t.Grid,1,false,"footer");
            AddText(s,64,679,970,25,d.source,14,t.Muted,false,"left","source",-1);
            AddText(s,1080,679,136,25,"Mello Charts",13,t.Muted,false,"right","brand",-1);
            return s;
        }
        static void SeriesLegend(Scene s,ExecutiveTheme t,bool baseline,string actual,string forecast,string secondColor=null) {
            double x=610;
            string[] labels=baseline?new[]{actual,forecast,"Baseline"}:new[]{actual,forecast};
            string[] colors={t.Palette.Actual,secondColor??t.ForecastLine,t.Palette.Baseline};
            for(int i=0;i<labels.Length;i++) {
                AddLine(s,x,204,x+22,204,colors[i],i==2?2:4,i==2,"legend");
                AddText(s,x+30,191,i==1?186:140,28,labels[i],15,t.Muted,false,"left","legend",-1);
                x+=i==1 && labels[i].Length>12?218:170;
            }
        }
        static void VerticalAxis(Scene s,ChartDocument d,ExecutiveTheme t,ExecutiveScale scale,double left,double right,double top,double bottom) {
            int count=(int)Math.Round((scale.High-scale.Low)/scale.Step);
            HashSet<double> values=new HashSet<double>();HashSet<string> labels=new HashSet<string>();
            for(int i=0;i<=count && i<20;i++) {
                double tick=scale.Low+i*scale.Step,value=Math.Abs(tick)<scale.Step*.001?0:tick*scale.Magnitude;
                string label=ExecutiveValue(d,value,true);
                if(!values.Add(value) || !labels.Add(label))continue;
                double y=scale.Map(value,bottom,top);
                if(value==0 || d.showGrid)AddLine(s,left,y,right,y,value==0?t.Axis:t.Grid,value==0?1.2:1,false,value==0?"zeroAxis":"grid");
                AddText(s,12,y-12,86,25,label,14,t.Muted,false,"right","tick",-1);
            }
        }
        static Scene BuildAdvancedScene(ChartDocument d) {
            ExecutiveTheme t=new ExecutiveTheme(d);Scene s=ExecutiveFrame(d,t);
            switch(d.type) {
                case "bars": HorizontalBars(s,d,t);break;
                case "lines":case "area":case "combo": TrendChart(s,d,t);break;
                case "stacked100": PercentColumns(s,d,t);break;
                case "pie":case "donut": CircularChart(s,d,t);break;
                case "pareto": ParetoChart(s,d,t);break;
                default:throw Bad("Tipo de gráfico inválido.");
            }
            return s;
        }
        static void Marker(Scene s,double x,double y,string color,ExecutiveTheme t,int category) {
            AddEllipse(s,x-5,y-5,10,10,t.Background,"markerHalo",category);
            AddEllipse(s,x-3.5,y-3.5,7,7,color,"marker",category);
        }
        static void TrendChart(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=112,right=1060,top=252,bottom=552;
            int n=d.rows.Count;double slot=(right-left)/n;
            List<double> values=d.rows.SelectMany(r=>d.showBaseline?new[]{r.actual,r.forecast,r.baseline}:new[]{r.actual,r.forecast}).ToList();
            ExecutiveScale scale=Scale(values);
            Func<double,double> y=v=>scale.Map(v,bottom,top);
            Func<int,double> x=i=>left+slot*(i+.5);
            SeriesLegend(s,t,d.showBaseline,"Realizado","Projetado");
            VerticalAxis(s,d,t,scale,left,right,top,bottom);
            if(d.type=="area") {
                for(int series=1;series>=0;series--) {
                    List<double> points=new List<double>{x(0),y(0)};
                    for(int i=0;i<n;i++){points.Add(x(i));points.Add(y(series==0?d.rows[i].actual:d.rows[i].forecast));}
                    points.Add(x(n-1));points.Add(y(0));
                    // A single observation has a column-width area rather than a zero-width polygon.
                    if(n==1)points=new List<double>{x(0)-slot*.22,y(0),x(0)-slot*.22,y(series==0?d.rows[0].actual:d.rows[0].forecast),x(0)+slot*.22,y(series==0?d.rows[0].actual:d.rows[0].forecast),x(0)+slot*.22,y(0)};
                    AddPolygon(s,points,Mix(t.Background,series==0?t.Palette.Actual:t.ForecastLine,.22),"area",series);
                }
            }
            if(d.showBaseline)for(int i=1;i<n;i++)AddLine(s,x(i-1),y(d.rows[i-1].baseline),x(i),y(d.rows[i].baseline),t.Palette.Baseline,2,true,"baseline");
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];
                if(d.type=="combo") {
                    double w=Math.Min(66,slot*.5);
                    AddRect(s,x(i)-w/2,Math.Min(y(row.actual),y(0)),w,Math.Abs(y(row.actual)-y(0)),t.Palette.Actual,"bar",i);
                    if(d.showLabels)AddText(s,x(i)-slot*.43,row.actual>=0?y(row.actual)-28:y(row.actual)+4,slot*.86,24,ExecutiveValue(d,row.actual,false),n>8?13:17,t.Ink,true,"center","value",i);
                } else if(i>0)AddLine(s,x(i-1),y(d.rows[i-1].actual),x(i),y(row.actual),t.Palette.Actual,3.5,false,"actualLine");
                if(i>0)AddLine(s,x(i-1),y(d.rows[i-1].forecast),x(i),y(row.forecast),t.ForecastLine,3,false,"forecastLine");
            }
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];
                if(d.type!="combo")Marker(s,x(i),y(row.actual),t.Palette.Actual,t,i);
                Marker(s,x(i),y(row.forecast),t.ForecastLine,t,i);
                if(d.showBaseline)AddEllipse(s,x(i)-2.5,y(row.baseline)-2.5,5,5,t.Palette.Baseline,"baseline",i);
                AddText(s,x(i)-slot*.46,574,slot*.92,42,Wrap(row.label,Math.Max(6,(int)(slot/10))),n>8?14:17,t.Muted,false,"center","category",i);
            }
            // End labels are spaced as a group, so equal or nearly equal values stay readable.
            List<Tuple<string,double,string>> ends=new List<Tuple<string,double,string>> {
                Tuple.Create("Realizado",d.rows[n-1].actual,t.Palette.Actual),
                Tuple.Create("Projetado",d.rows[n-1].forecast,t.ForecastLine)
            };
            if(d.showBaseline)ends.Add(Tuple.Create("Baseline",d.rows[n-1].baseline,t.Palette.Baseline));
            ends=ends.OrderBy(e=>y(e.Item2)).ToList();
            double[] positions=ends.Select(e=>Math.Max(top,Math.Min(bottom-40,y(e.Item2)-20))).ToArray();
            for(int i=1;i<positions.Length;i++)positions[i]=Math.Max(positions[i],positions[i-1]+50);
            if(positions[positions.Length-1]>bottom-40) {
                positions[positions.Length-1]=bottom-40;
                for(int i=positions.Length-2;i>=0;i--)positions[i]=Math.Min(positions[i],positions[i+1]-50);
            }
            for(int i=0;i<ends.Count;i++) {
                var e=ends[i];AddLine(s,x(n-1)+7,y(e.Item2),1080,positions[i]+19,e.Item3,1.3,false,"seriesLeader");
                AddText(s,1088,positions[i],128,44,e.Item1+(d.showLabels?"\n"+ExecutiveValue(d,e.Item2,false):""),16,ColorText(d,e.Item3,t.Background),true,"left","seriesLabel",-1);
            }
            if(d.rows[0].actual>0 && n>1) {
                double change=(d.rows[n-1].actual/d.rows[0].actual-1)*100;
                if(!Double.IsInfinity(change) && !Double.IsNaN(change))
                    AddText(s,112,623,1104,25,"Variação do realizado entre o primeiro e o último ponto: "+(change>0?"+":"")+FormatNumber(change,1)+"%",16,t.Ink,true,"left","insight",-1);
            }
        }
        static void HorizontalBars(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=330,right=1100,top=252,bottom=608;
            ExecutiveScale scale=Scale(d.rows.SelectMany(r=>d.showBaseline?new[]{r.actual,r.forecast,r.baseline}:new[]{r.actual,r.forecast}));
            Func<double,double> x=v=>scale.Map(v,left,right);int n=d.rows.Count;
            double slot=(bottom-top)/n,h=Math.Min(18,slot*.24);
            SeriesLegend(s,t,d.showBaseline,"Realizado","Projetado",t.Palette.Forecast);
            if(d.style!="minimal")for(int i=0;i<n;i+=2) {
                double cy=top+slot*(i+.5);
                AddRect(s,64,cy-slot*.43,1152,slot*.86,t.Band,"rowBand",i);
            }
            int ticks=(int)Math.Round((scale.High-scale.Low)/scale.Step);
            HashSet<double> values=new HashSet<double>();HashSet<string> labels=new HashSet<string>();
            for(int i=0;i<=ticks && i<20;i++) {
                double tick=scale.Low+i*scale.Step,v=Math.Abs(tick)<scale.Step*.001?0:tick*scale.Magnitude;
                string label=ExecutiveValue(d,v,true);
                if(!values.Add(v) || !labels.Add(label))continue;
                double xx=x(v);
                if(v==0 || d.showGrid)AddLine(s,xx,top-10,xx,bottom,v==0?t.Axis:t.Grid,v==0?1.4:1,false,v==0?"zeroAxis":"grid");
                AddText(s,xx-45,top-37,90,22,label,13,t.Muted,false,"center","tick",-1);
            }
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];double cy=top+slot*(i+.5);
                AddText(s,64,cy-slot*.36,230,slot*.72,Wrap(row.label,24),n>8?13:17,t.Ink,true,"left","category",i);
                for(int j=0;j<2;j++) {
                    double value=j==0?row.actual:row.forecast,yy=cy+(j==0?-h-2:2);
                    AddRect(s,Math.Min(x(0),x(value)),yy,Math.Abs(x(value)-x(0)),h,j==0?t.Palette.Actual:t.Palette.Forecast,"bar",i);
                    if(d.showLabels && (n<=8 || j==0)) {
                        double width=Math.Abs(x(value)-x(0));
                        bool inside=width>150;
                        double labelX=inside?(value>=0?x(value)-100:x(value)+6):(value>=0?x(value)+7:x(value)-100);
                        if(d.showBaseline) {
                            double baselineX=x(row.baseline);
                            if(baselineX>=labelX-5 && baselineX<=labelX+99) {
                                inside=false;
                                labelX=value>=0?Math.Max(x(value)+7,baselineX+7):Math.Min(x(value)-100,baselineX-101);
                            }
                        }
                        AddText(s,labelX,yy-2,94,Math.Max(h+4,17),ExecutiveValue(d,value,false),n>8?12:16,
                            inside?FillInk(d,j==0?t.Palette.Actual:t.Palette.Forecast,j==0?(d.style=="dark"?"#101E2E":"#FFFFFF"):(d.style=="dark"?"#FFFFFF":"#173E45")):t.Ink,true,value>=0?"right":"left","value",i);
                    }
                }
                if(d.showBaseline)AddLine(s,x(row.baseline),cy-h-6,x(row.baseline),cy+h+6,t.Palette.Baseline,2.5,true,"baseline");
            }
            AddText(s,64,625,1152,25,"Comparação por categoria · as duas séries não são somadas.",15,t.Muted,false,"left","insight",-1);
        }
        static void PercentColumns(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=112,right=1216,top=252,bottom=552;
            int n=d.rows.Count;double slot=(right-left)/n,w=Math.Min(104,slot*.62);
            SeriesLegend(s,t,false,"Realizado","Projetado restante",t.Palette.Forecast);
            for(int p=0;p<=100;p+=25) {
                double y=bottom-p/100.0*(bottom-top);
                if(p==0 || d.showGrid)AddLine(s,left,y,right,y,p==0?t.Axis:t.Grid,1,false,p==0?"zeroAxis":"grid");
                AddText(s,12,y-12,86,24,p+"%",14,t.Muted,false,"right","tick",-1);
            }
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];double cx=left+slot*(i+.5),sum=row.actual+row.forecast;
                double fraction=sum==0?0:row.actual/sum,actualHeight=(bottom-top)*fraction,forecastHeight=(bottom-top)*(1-fraction);
                if(sum>0) {
                    AddRect(s,cx-w/2,bottom-actualHeight,w,actualHeight,t.Palette.Actual,"bar",i);
                    AddRect(s,cx-w/2,top,w,forecastHeight,t.Palette.Forecast,"bar",i);
                    if(d.showLabels) {
                        if(actualHeight>40)AddText(s,cx-w/2+4,bottom-actualHeight/2-13,w-8,26,FormatNumber(fraction*100,0)+"%",17,FillInk(d,t.Palette.Actual,d.style=="dark"?"#101E2E":"#FFFFFF"),true,"center","percentValue",i);
                        if(forecastHeight>40)AddText(s,cx-w/2+4,top+forecastHeight/2-13,w-8,26,FormatNumber((1-fraction)*100,0)+"%",17,FillInk(d,t.Palette.Forecast,d.style=="dark"?"#FFFFFF":"#173E45"),true,"center","percentValue",i);
                    }
                } else AddText(s,cx-slot*.46,bottom-28,slot*.92,24,"Sem base",14,t.Muted,false,"center","empty",i);
                AddText(s,cx-slot*.46,574,slot*.92,48,Wrap(row.label,Math.Max(6,(int)(slot/10))),n>8?14:17,t.Muted,false,"center","category",i);
            }
            AddText(s,112,628,1104,25,"Realizado + projetado restante = 100% da composição de cada categoria.",15,t.Muted,false,"left","insight",-1);
        }
        static string[] BuiltinSliceColors(ChartDocument d) {
            string[] colors=d.palette=="forest"?new[]{"#286153","#66A899","#A7CBB1","#A86D40","#D6B784","#657F83"}:
                d.palette=="ink"?new[]{"#3C4A63","#6E85A8","#A2B5D1","#A97869","#C7B3AC","#839FA2"}:
                new[]{"#244F73","#3A8497","#72B9AF","#AFCEB5","#B2763E","#D4B58C"};
            if(d.style=="dark")colors=colors.Select(c=>Mix("#F3F6FB",c,.74)).ToArray();
            return colors.Concat(colors.Select(c=>Mix(d.style=="dark"?"#101E2E":"#FFFFFF",c,.62))).ToArray();
        }
        static string[] SliceColors(ChartDocument d) {
            string[] colors=BuiltinSliceColors(d);
            if(d.customColors!=null&&d.customColors.categories!=null)for(int i=0;i<d.customColors.categories.Count;i++)colors[i]=d.customColors.categories[i];
            return colors;
        }
        static void CircularChart(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double cx=340,cy=430,radius=172;double total=d.rows.Sum(r=>r.value),angle=-Math.PI/2;
            bool donut=d.type=="donut";double inner=donut?radius*.64:0;string[] colors=SliceColors(d);
            int positives=d.rows.Count(r=>r.value>0),n=d.rows.Count;double rowHeight=Math.Min(48,350.0/n);
            AddText(s,710,229,300,22,"CATEGORIA",12,t.Muted,true,"left","legend",-1);
            if(d.showLabels) {
                AddText(s,1010,229,90,22,"PARTICIPAÇÃO",11,t.Muted,true,"right","legend",-1);
                AddText(s,1106,229,110,22,"VALOR",11,t.Muted,true,"right","legend",-1);
            }
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];double sweep=row.value/total*Math.PI*2;
                if(sweep>0) {
                    double gap=positives>1?Math.Min(.018,sweep*.08):0,start=angle+gap/2,end=angle+sweep-gap/2;
                    int steps=Math.Max(2,(int)Math.Ceiling((end-start)/(.025)));
                    List<double> points=new List<double>();
                    if(!donut) {points.Add(cx);points.Add(cy);}
                    for(int k=0;k<=steps;k++) {double a=start+(end-start)*k/steps;points.Add(cx+radius*Math.Cos(a));points.Add(cy+radius*Math.Sin(a));}
                    if(donut)for(int k=steps;k>=0;k--) {double a=start+(end-start)*k/steps;points.Add(cx+inner*Math.Cos(a));points.Add(cy+inner*Math.Sin(a));}
                    AddPolygon(s,points,colors[i],"slice",i);
                }
                angle+=sweep;
                double yy=270+i*rowHeight;
                if(i==0 && d.style!="minimal")AddRect(s,690,yy-7,526,rowHeight-1,t.Band,"legendBand",i);
                AddRect(s,710,yy+5,12,12,colors[i],"legendSwatch",i);
                AddText(s,735,yy-1,260,rowHeight-3,row.label,17,t.Ink,i==0,"left","category",i);
                if(d.showLabels) {
                    double share=row.value/total*100;
                    string shareLabel=row.value>0 && share<.1?"<0,1%":FormatNumber(share,1)+"%";
                    AddText(s,1005,yy-1,95,rowHeight-3,shareLabel,17,t.Ink,true,"right","percentValue",i);
                    AddText(s,1106,yy-1,110,rowHeight-3,ExecutiveValue(d,row.value,false),16,t.Muted,false,"right","value",i);
                }
            }
            if(donut) {
                AddText(s,cx-115,cy-39,230,54,d.showLabels?ExecutiveValue(d,total,false):"Composição",d.showLabels?35:25,t.Ink,true,"center",d.showLabels?"centerValue":"centerCaption",-1);
                AddText(s,cx-115,cy+19,230,28,"TOTAL · "+d.unit,14,t.Muted,false,"center","centerCaption",-1);
            }
            int largest=Enumerable.Range(0,n).OrderByDescending(i=>d.rows[i].value).First();
            AddText(s,64,629,1152,25,"Maior participação: "+d.rows[largest].label+" · "+FormatNumber(d.rows[largest].value/total*100,1)+"% do total.",16,t.Ink,true,"left","insight",-1);
        }
        static void ParetoChart(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=112,right=1110,top=252,bottom=552;
            var rows=d.rows.Select((row,index)=>new{Row=row,Index=index}).OrderByDescending(r=>r.Row.value).ToList();
            double total=rows.Sum(r=>r.Row.value),running=0,slot=(right-left)/rows.Count,w=Math.Min(98,slot*.62);
            ExecutiveScale scale=Scale(rows.Select(r=>r.Row.value));
            Func<double,double> y=v=>scale.Map(v,bottom,top);
            SeriesLegend(s,t,false,"Valor","Acumulado (%)",t.Palette.Baseline);
            VerticalAxis(s,d,t,scale,left,right,top,bottom);
            AddLine(s,left,bottom-.8*(bottom-top),right,bottom-.8*(bottom-top),t.Axis,1.5,true,"threshold");
            for(int p=0;p<=100;p+=20)AddText(s,1128,bottom-p/100.0*(bottom-top)-12,88,25,p+"%",14,t.Muted,false,"right","percentTick",-1);
            double previousX=0,previousY=0,covered=0;int priorityCount=0;bool reached=false;
            for(int i=0;i<rows.Count;i++) {
                var row=rows[i];double cx=left+slot*(i+.5),before=running/total;running+=row.Row.value;
                if(!reached) {priorityCount++;covered=running/total;if(covered>=.8)reached=true;}
                AddRect(s,cx-w/2,y(row.Row.value),w,y(0)-y(row.Row.value),before<.8?t.Palette.Actual:t.Palette.Forecast,"bar",row.Index);
                if(d.showLabels)AddText(s,cx-slot*.46,y(row.Row.value)-29,slot*.92,25,ExecutiveValue(d,row.Row.value,false),rows.Count>8?14:18,t.Ink,true,"center","value",row.Index);
                double yy=bottom-running/total*(bottom-top);
                if(i>0)AddLine(s,previousX,previousY,cx,yy,t.Palette.Baseline,3,false,"cumulative");
                Marker(s,cx,yy,t.Palette.Baseline,t,row.Index);previousX=cx;previousY=yy;
                AddText(s,cx-slot*.46,574,slot*.92,47,Wrap(row.Row.label,Math.Max(6,(int)(slot/10))),rows.Count>8?14:17,t.Muted,false,"center","category",row.Index);
            }
            AddText(s,112,627,1104,25,priorityCount+" de "+rows.Count+" categorias representam "+FormatNumber(covered*100,1)+"% do total. Referência tracejada: 80%.",16,t.Ink,true,"left","insight",-1);
        }
    }
}
