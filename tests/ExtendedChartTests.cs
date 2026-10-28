using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using MelloCharts;

public static class ExtendedChartTests {
    static int passed, failed;
    static readonly string[] Added = {"bars","lines","area","combo","stacked100","pie","donut","pareto"};
    static void Is(bool value, string message) { if (!value) throw new Exception(message); }
    static void Test(string name, Action action) {
        try { action(); passed++; Console.WriteLine("PASS " + name); }
        catch (Exception ex) { failed++; Console.Error.WriteLine("FAIL " + name + ": " + ex.Message); }
    }
    static void Reject(Action action) { bool rejected=false; try { action(); } catch(ArgumentException) { rejected=true; } Is(rejected,"Expected a validation error"); }
    static bool ValueType(string type) { return type=="pie" || type=="donut" || type=="pareto"; }
    [STAThread] public static int Main() {
        foreach (string current in Added) {
            string type=current;
            Test(type+" has a usable fictional example", delegate {
                ChartDocument d=ChartCore.Validate(ChartCore.Example(type));
                Is(d.type==type && d.rows.Count>=4,"Wrong example");
                Is(d.source.ToLowerInvariant().Contains("fict"),"Example must be fictional");
            });
            Test(type+" JSON keeps data and chart identity", delegate {
                ChartDocument d=ChartCore.Example(type); d.style="dark"; d.compactNumbers=true;
                ChartDocument copy=ChartCore.Deserialize(ChartCore.Serialize(d));
                Is(ChartCore.Serialize(copy)==ChartCore.Serialize(d),"JSON changed the document");
            });
            Test(type+" CSV roundtrip uses the right columns", delegate {
                ChartDocument d=ChartCore.Example(type);
                string csv=ChartCore.ToCsv(d);
                Is(csv.Contains(ValueType(type)?"categoria;valor":"categoria;realizado;projetado;baseline"),"Wrong header");
                ChartDocument copy=ChartCore.ImportDelimited(csv,type);
                Is(copy.rows.Count==d.rows.Count,"Rows lost");
                for(int i=0;i<d.rows.Count;i++) Is(ValueType(type)?copy.rows[i].value==d.rows[i].value:copy.rows[i].actual==d.rows[i].actual,"Value changed");
            });
        }
        Test("single-value input accepts PTBR decimals without dummy series", delegate {
            foreach(string type in new[]{"pie","donut","pareto"}) {
                ChartDocument d=ChartCore.ImportDelimited("categoria;valor\nEngenharia;1.234,56\nEquipamentos;765,44",type);
                Is(d.rows[0].value==1234.56 && d.rows[1].value==765.44,"Wrong single values");
            }
        });
        Test("composition charts reject negative and all-zero datasets", delegate {
            foreach(string type in new[]{"pie","donut","pareto"}) {
                ChartDocument d=ChartCore.Example(type); d.rows[0].value=-1;
                Reject(delegate { ChartCore.Validate(d); });
                foreach(ChartRow row in d.rows) row.value=0;
                Reject(delegate { ChartCore.Validate(d); });
            }
        });
        Test("100-percent composition rejects negative portions", delegate {
            ChartDocument d=ChartCore.Example("stacked100"); d.rows[0].forecast=-1;
            Reject(delegate { ChartCore.Validate(d); });
        });
        Test("legacy three chart types retain their models", delegate {
            foreach(string type in new[]{"buildup","waterfall","columns"}) {
                ChartDocument d=ChartCore.Example(type);
                Is(ChartCore.Serialize(ChartCore.Deserialize(ChartCore.Serialize(d)))==ChartCore.Serialize(d),"Legacy changed");
            }
        });
        Test("horizontal chart uses horizontal bars", delegate {
            var bars=ChartCore.BuildScene(ChartCore.Example("bars")).Elements.Where(e=>e.Role=="bar").ToList();
            Is(bars.Count==8 && bars.All(e=>e.W>e.H*2),"Bars are not horizontal");
        });
        Test("line chart connects both series with editable markers", delegate {
            Scene s=ChartCore.BuildScene(ChartCore.Example("lines"));
            Is(s.Elements.Any(e=>e.Role=="actualLine") && s.Elements.Any(e=>e.Role=="forecastLine"),"Missing line series");
            Is(s.Elements.Count(e=>e.Type=="ellipse" && e.Role=="marker")>=12,"Missing point markers");
            Is(!s.Elements.Any(e=>e.Role=="bar"),"Unexpected columns");
        });
        Test("area chart includes two independent filled regions", delegate {
            Scene s=ChartCore.BuildScene(ChartCore.Example("area"));
            Is(s.Elements.Count(e=>e.Type=="polygon" && e.Role=="area")==2,"Missing area polygons");
        });
        Test("combination uses one bar series and one projected line", delegate {
            ChartDocument d=ChartCore.Example("combo"); Scene s=ChartCore.BuildScene(d);
            Is(s.Elements.Count(e=>e.Role=="bar")==d.rows.Count,"Wrong column count");
            Is(s.Elements.Any(e=>e.Role=="forecastLine"),"Missing projected line");
        });
        Test("100-percent segments share a column and an equal total height", delegate {
            ChartDocument d=ChartCore.Example("stacked100"); Scene s=ChartCore.BuildScene(d);
            double expected=-1;
            for(int i=0;i<d.rows.Count;i++) {
                var bars=s.Elements.Where(e=>e.Role=="bar" && e.Category==i).ToList();
                Is(bars.Count==2 && bars[0].X==bars[1].X,"Segments are not complementary");
                double h=bars.Sum(e=>e.H); if(expected<0) expected=h;
                Is(Math.Abs(h-expected)<1e-7,"Column does not sum to100%");
            }
            d.rows[0].actual=0;d.rows[0].forecast=0;
            Is(ChartCore.BuildScene(d).Elements.Any(e=>e.Text=="Sem base"),"Zero total treated as100%");
        });
        Test("pie and donut use one editable slice per positive category", delegate {
            foreach(string type in new[]{"pie","donut"}) {
                ChartDocument d=ChartCore.Example(type); Scene s=ChartCore.BuildScene(d);
                Is(s.Elements.Count(e=>e.Type=="polygon" && e.Role=="slice")==d.rows.Count,"Wrong slice count");
                Is(s.Elements.Any(e=>e.Role=="centerValue")==(type=="donut"),"Wrong center label");
                d.rows[0].value=0;
                Is(ChartCore.BuildScene(d).Elements.Count(e=>e.Role=="slice")==d.rows.Count-1,"Zero slice rendered");
            }
        });
        Test("Pareto sorts only its drawing and reaches100 percent", delegate {
            ChartDocument d=ChartCore.Example("pareto");d.rows.Reverse();string before=ChartCore.Serialize(d);
            Scene s=ChartCore.BuildScene(d);
            Is(ChartCore.Serialize(d)==before,"Drawing reordered the source data");
            var bars=s.Elements.Where(e=>e.Role=="bar").OrderBy(e=>e.X).ToList();
            Is(bars[0].H>=bars.Last().H && bars[0].Category==d.rows.Count-1,"Not sorted by descending value");
            var cumulative=s.Elements.Where(e=>e.Role=="cumulative" && e.Type=="line").ToList();
            Is(cumulative.Count==d.rows.Count-1 && Math.Abs(cumulative.Last().Y2-252)<1e-7,"Cumulative does not reach100%");
            Is(s.Elements.Any(e=>e.Role=="threshold"),"Missing80% reference");
        });
        Test("all styles and palettes keep twelve-category scenes inside the slide", delegate {
            foreach(string type in Added) foreach(string style in new[]{"executive","minimal","dark"}) foreach(string palette in new[]{"ocean","forest","ink"}) {
                ChartDocument d=ChartCore.Example(type);d.style=style;d.palette=palette;d.title=new string('W',80);d.rows.Clear();
                for(int i=0;i<12;i++)d.rows.Add(new ChartRow{label="Categoria longa de comparação "+i,actual=1000000000,forecast=i%3==0?0:1000000000,baseline=1000000000,value=i+1});
                Scene s=ChartCore.BuildScene(d);
                foreach(SceneElement e in s.Elements) {
                    Is(!Double.IsNaN(e.X) && !Double.IsNaN(e.Y) && !Double.IsInfinity(e.Y),"Nonfinite scene "+type);
                    double maxX=e.Type=="line"?Math.Max(e.X,e.X2):e.X+e.W,maxY=e.Type=="line"?Math.Max(e.Y,e.Y2):e.Y+e.H;
                    Is(e.X>=-0.01 && e.Y>=-0.01 && maxX<=1280.01 && maxY<=720.01,"Outside slide "+type+" "+e.Role);
                }
            }
        });
        Test("zero and subnormal comparative inputs keep finite coordinates", delegate {
            foreach(string type in new[]{"bars","lines","area","combo","stacked100"}) {
                ChartDocument d=ChartCore.Example(type);
                foreach(ChartRow row in d.rows)row.actual=row.forecast=row.baseline=0;
                foreach(double value in new[]{0,Double.Epsilon}) {
                    d.rows[0].actual=value;
                    foreach(SceneElement e in ChartCore.BuildScene(d).Elements)Is(!Double.IsNaN(e.X+e.Y+e.W+e.H),"Nonfinite zero scene "+type);
                }
            }
        });
        Test("hidden data labels retain the named categories", delegate {
            foreach(string type in Added) {
                ChartDocument d=ChartCore.Example(type);d.showLabels=false;Scene s=ChartCore.BuildScene(d);
                Is(!s.Elements.Any(e=>e.Role=="value" || e.Role=="centerValue" || e.Role=="percentValue"),"Value label still visible "+type);
                Is(s.Elements.Any(e=>e.Role=="category"),"Categories disappeared "+type);
            }
        });
        Test("negative comparison values and single observations remain finite", delegate {
            foreach(string type in Added) {
                ChartDocument d=ChartCore.Example(type);d.rows=d.rows.Take(1).ToList();
                if(type=="bars" || type=="lines" || type=="combo") { d.rows[0].actual=-1000000000;d.rows[0].forecast=-25;d.rows[0].baseline=100; }
                foreach(SceneElement e in ChartCore.BuildScene(d).Elements)Is(!Double.IsNaN(e.X+e.Y+e.W+e.H),"Invalid single observation "+type);
            }
        });
        Test("new chart text fits long-title and twelve-category frames", delegate {
            using(Bitmap bitmap=new Bitmap(1280,720))using(Graphics graphics=Graphics.FromImage(bitmap)) {
                foreach(string type in Added) foreach(string style in new[]{"executive","minimal","dark"}) {
                    ChartDocument d=ChartCore.Example(type);d.style=style;d.title=new string('W',80);d.rows.Clear();
                    for(int i=0;i<12;i++)d.rows.Add(new ChartRow{label="Categoria longa de comparação "+i,actual=1000000000,forecast=1000000000,baseline=1000000000,value=1000000000});
                    foreach(SceneElement e in ChartCore.BuildScene(d).Elements.Where(e=>e.Type=="text"))
                        using(Font font=new Font("Arial",(float)e.Size,e.Bold?FontStyle.Bold:FontStyle.Regular,GraphicsUnit.Pixel))
                        using(StringFormat format=new StringFormat()) {
                            SizeF measured=graphics.MeasureString(e.Text,font,new SizeF((float)e.W,10000),format);
                            Is(measured.Height<=e.H+.6,"Clipped text in "+type+"/"+e.Role+": "+measured.Height+" > "+e.H);
                        }
                }
            }
        });
        Test("new polygons contain closed finite editable coordinates", delegate {
            foreach(string type in new[]{"area","pie","donut"}) {
                foreach(SceneElement e in ChartCore.BuildScene(ChartCore.Example(type)).Elements.Where(e=>e.Type=="polygon")) {
                    Is(e.Points!=null && e.Points.Length>=8 && e.Points.Length%2==0,"Missing vertices");
                    Is(e.Points[0]==e.Points[e.Points.Length-2] && e.Points[1]==e.Points[e.Points.Length-1],"Open polygon");
                    Is(e.Points.All(v=>!Double.IsInfinity(v)&&!Double.IsNaN(v)),"Nonfinite vertex");
                }
            }
        });
        Test("legends exactly match the corresponding plotted series", delegate {
            foreach(string type in new[]{"bars","stacked100","pareto"}) {
                Scene s=ChartCore.BuildScene(ChartCore.Example(type));
                var legend=s.Elements.Where(e=>e.Role=="legend" && e.Type=="line").ToList();
                string seriesColor=type=="pareto"?s.Elements.First(e=>e.Role=="cumulative").Color:s.Elements.Where(e=>e.Role=="bar").Skip(1).First().Color;
                Is(legend[1].Color==seriesColor,"Legend color differs from data in "+type);
            }
        });
        Test("horizontal baseline ticks avoid value label frames", delegate {
            Scene s=ChartCore.BuildScene(ChartCore.Example("bars"));
            foreach(SceneElement tick in s.Elements.Where(e=>e.Role=="baseline"))
                foreach(SceneElement label in s.Elements.Where(e=>e.Role=="value"))
                    Is(!(tick.X>=label.X-2 && tick.X<=label.X+label.W+2 && Math.Max(tick.Y,tick.Y2)>=label.Y && Math.Min(tick.Y,tick.Y2)<=label.Y+label.H),"Tick obscures a numeric label");
        });
        Test("subnormal numeric axes show unique labels at correct positions", delegate {
            foreach(string type in new[]{"bars","lines"}) {
                ChartDocument d=ChartCore.Example(type);foreach(ChartRow row in d.rows)row.actual=row.forecast=row.baseline=Double.Epsilon;
                Scene s=ChartCore.BuildScene(d);var ticks=s.Elements.Where(e=>e.Role=="tick").Select(e=>e.Text).ToList();
                Is(ticks.Count==ticks.Distinct().Count(),"Duplicate axis labels "+type);
                Is(s.Elements.Count(e=>e.Role=="zeroAxis")==1,"Zero reference is repeated "+type);
            }
        });
        Test("negligible circular slices retain data without degenerate native polygons", delegate {
            foreach(string type in new[]{"pie","donut"}) {
                ChartDocument d=ChartCore.Example(type);
                d.rows=new List<ChartRow>{new ChartRow{label="Principal",value=1000000000},new ChartRow{label="Residual",value=1}};
                string before=ChartCore.Serialize(d);Scene s=ChartCore.BuildScene(d);
                foreach(SceneElement e in s.Elements.Where(e=>e.Type=="polygon")) {
                    var vertices=new HashSet<string>();
                    for(int i=0;i<e.Points.Length;i+=2)vertices.Add(((float)(24+e.Points[i]*.5)).ToString("R")+","+((float)(24+e.Points[i+1]*.5)).ToString("R"));
                    Is(vertices.Count>=3,"Polygon collapses at native float precision");
                }
                Is(ChartCore.Serialize(d)==before,"Small value discarded from data");
                Is(s.Elements.Any(e=>e.Role=="percentValue" && e.Category==1 && e.Text=="<0,1%"),"Small positive share presented as zero");
            }
        });
        Console.WriteLine(passed+" extended model checks passed; "+failed+" failed");
        return failed==0?0:1;
    }
}
