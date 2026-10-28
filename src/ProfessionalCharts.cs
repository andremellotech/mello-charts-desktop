using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MelloCharts {
    public static partial class ChartCore {
        static bool IsProfessionalType(string type) { return type=="gantt" || type=="scatter" || type=="bubble"; }
        static ChartDocument ProfessionalExample(string type) {
            ChartDocument d=new ChartDocument {version=1,type=type,title=type=="gantt"?"Cronograma de implantação":type=="scatter"?"Relação entre esforço e resultado":"Oportunidades por esforço e retorno",subtitle=type=="gantt"?"Atividades planejadas e percentual concluído por etapa":type=="scatter"?"Compare esforço (X) e retorno (Y) de cada iniciativa":"Esforço (X) e retorno (Y); área representa o porte da iniciativa",source="Exemplo com dados fictícios",unit=type=="gantt"?"Datas inclusivas":"Índice",palette="ocean",style="executive",decimals=0,showLabels=true,showBaseline=false,includeTotal=false,showGrid=true,compactNumbers=false,rows=new List<ChartRow>()};
            string[] labels={"Diagnóstico","Planejamento","Desenvolvimento","Validação","Implantação"};
            for(int i=0;i<labels.Length;i++) d.rows.Add(type=="gantt"?new ChartRow {label=labels[i],startDate=new DateTime(2026,10,5).AddDays(i*9).ToString("yyyy-MM-dd",CultureInfo.InvariantCulture),endDate=new DateTime(2026,10,5).AddDays(i*9+12).ToString("yyyy-MM-dd",CultureInfo.InvariantCulture),progress=new[]{100.0,80,45,10,0}[i]}:new ChartRow {label=labels[i],x=new[]{-20.0,15,35,65,80}[i],y=new[]{30.0,-10,55,75,40}[i],size=type=="bubble"?new[]{20.0,35,55,80,45}[i]:0});
            return d;
        }
        static DateTime ProfessionalDate(string value,bool allowImport) {
            DateTime date;
            if(value!=null && DateTime.TryParseExact(value,allowImport?new[]{"yyyy-MM-dd","dd/MM/yyyy","d/M/yyyy"}:new[]{"yyyy-MM-dd"},CultureInfo.InvariantCulture,DateTimeStyles.None,out date))return date;
            double serial;
            if(allowImport && value!=null && Double.TryParse(value,NumberStyles.AllowDecimalPoint,CultureInfo.InvariantCulture,out serial) && serial>=1 && serial<2958466 && Math.Truncate(serial)!=60) {
                try {return DateTime.FromOADate(serial<60?serial+1:serial).Date;}catch(ArgumentException){}
            }
            throw Bad("Data inválida. Use AAAA-MM-DD ou DD/MM/AAAA; revise o calendário.");
        }
        static void ValidateProfessionalRow(ChartRow r,string type) {
            Number(r.x);Number(r.y);Number(r.size);Number(r.progress);
            if(type=="gantt") {
                DateTime start=ProfessionalDate(r.startDate,false),end=ProfessionalDate(r.endDate,false);
                if(end<start)throw Bad("A data final deve ser igual ou posterior à data inicial.");
                if(r.progress<0 || r.progress>100)throw Bad("O progresso deve estar entre 0 e 100%.");
            } else if(type=="bubble" && r.size<=0)throw Bad("O tamanho de cada bolha deve ser positivo.");
        }
        static void ReadProfessionalRow(IDictionary<string,object> raw,ChartRow row,string type) {
            if(type=="gantt") {row.startDate=StringField(raw,"startDate");row.endDate=StringField(raw,"endDate");row.progress=NumericField(raw,"progress");}
            else {row.x=NumericField(raw,"x");row.y=NumericField(raw,"y");if(type=="bubble")row.size=NumericField(raw,"size");}
            foreach(string key in new[]{"x","y","size","progress"})if(raw.ContainsKey(key)) {double n=NumericField(raw,key);switch(key){case "x":row.x=n;break;case "y":row.y=n;break;case "size":row.size=n;break;case "progress":row.progress=n;break;}}
        }
        static string ProfessionalHeader(string text) {
            switch(Norm(text)) {
                case "categoria":case "category":case "atividade":case "etapa":case "projeto":case "label":return "label";
                case "inicio":case "data inicial":case "startdate":return "startDate";
                case "fim":case "data final":case "enddate":return "endDate";
                case "progresso":case "progress":return "progress";
                case "tamanho":case "size":return "size";
                default:return Norm(text);
            }
        }
        static double ProfessionalNumber(string text,bool invariantNumbers) {
            if(!invariantNumbers)return ParseNumber(text);
            double n;if(String.IsNullOrWhiteSpace(text) || !Double.TryParse(text,NumberStyles.Float,CultureInfo.InvariantCulture,out n))throw Bad("Informe um valor numérico válido.");Number(n);return n;
        }
        static ChartDocument ImportProfessionalRows(List<List<string>> table,string type,bool invariantNumbers) {
            if(table==null || table.Count<2)throw Bad("Inclua cabeçalho e pelo menos uma linha.");
            if(table.Count-1>RowLimit(type))throw Bad("Use de 1 a "+RowLimit(type)+" categorias.");
            string[] headers=table[0].Select(ProfessionalHeader).ToArray();
            if(headers.Length>64 || headers.Any(String.IsNullOrEmpty) || headers.Distinct().Count()!=headers.Length)throw Bad("Há cabeçalhos vazios ou duplicados.");
            string[] required=type=="gantt"?new[]{"label","startDate","endDate","progress"}:type=="bubble"?new[]{"label","x","y","size"}:new[]{"label","x","y"};
            if(required.Any(key=>!headers.Contains(key)))throw Bad("Cabeçalho esperado: "+ChartCatalog.Get(type).Header+".");
            ChartDocument d=ProfessionalExample(type);
            d.rows.Clear();
            for(int i=1;i<table.Count;i++) {
                if(table[i].Count!=headers.Length)throw Bad("Linha "+(i+1)+": quantidade de colunas diferente do cabeçalho.");
                ChartRow row=new ChartRow();
                try {foreach(string key in required) {
                    string text=table[i][Array.IndexOf(headers,key)];
                    switch(key) {
                        case "label":row.label=Regex.Replace(text.Trim(),"\\s+"," ");break;
                        case "startDate":row.startDate=ProfessionalDate(text.Trim(),true).ToString("yyyy-MM-dd",CultureInfo.InvariantCulture);break;
                        case "endDate":row.endDate=ProfessionalDate(text.Trim(),true).ToString("yyyy-MM-dd",CultureInfo.InvariantCulture);break;
                        case "progress":row.progress=ProfessionalNumber(text,invariantNumbers);break;
                        case "x":row.x=ProfessionalNumber(text,invariantNumbers);break;
                        case "y":row.y=ProfessionalNumber(text,invariantNumbers);break;
                        case "size":row.size=ProfessionalNumber(text,invariantNumbers);break;
                    }
                }ValidateProfessionalRow(row,type);}catch(ArgumentException e){throw Bad("Linha "+(i+1)+": "+e.Message);}
                d.rows.Add(row);
            }
            return Validate(d);
        }
        static string ProfessionalCsv(ChartDocument d) {
            StringBuilder csv=new StringBuilder("\uFEFF"+ChartCatalog.Get(d.type).Header);
            foreach(ChartRow row in d.rows) {
                csv.Append("\r\n").Append(CsvLabel(row.label)).Append(';');
                if(d.type=="gantt")csv.Append(row.startDate).Append(';').Append(row.endDate).Append(';').Append(CsvNumber(row.progress));
                else {csv.Append(CsvNumber(row.x)).Append(';').Append(CsvNumber(row.y));if(d.type=="bubble")csv.Append(';').Append(CsvNumber(row.size));}
            }
            return csv.ToString();
        }
        static Scene BuildProfessionalScene(ChartDocument d) {
            ExecutiveTheme t=new ExecutiveTheme(d);Scene scene=ExecutiveFrame(d,t);
            if(d.type=="gantt")ProfessionalGantt(scene,d,t);else ProfessionalPoints(scene,d,t);
            return scene;
        }
        static string ProfessionalColor(ChartDocument d,ExecutiveTheme t,int index) {
            if(d.customColors!=null && d.customColors.categories!=null && d.customColors.categories.Count>0)return d.customColors.categories[index%d.customColors.categories.Count];
            return t.Palette.Actual;
        }
        static void ProfessionalGantt(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=344,right=1138,top=264,bottom=617;
            DateTime origin=d.rows.Min(r=>ProfessionalDate(r.startDate,false)),last=d.rows.Max(r=>ProfessionalDate(r.endDate,false));
            double span=(last-origin).TotalDays+1,slot=(bottom-top)/d.rows.Count;
            Func<DateTime,double> x=date=>left+(date-origin).TotalDays/span*(right-left);
            AddText(s,64,228,255,22,"ATIVIDADE",12,t.Muted,true,"left","legend",-1);
            AddText(s,1148,228,68,22,"AVANÇO",11,t.Muted,true,"right","legend",-1);
            if(d.style!="minimal")for(int i=0;i<d.rows.Count;i+=2)AddRect(s,64,top+slot*(i+.02),1152,slot*.96,t.Band,"rowBand",i);
            for(int i=0;i<=5;i++) {
                double days=Math.Floor((span-1)*i/5),xx=left+days/span*(right-left);
                if(i>0 && days==Math.Floor((span-1)*(i-1)/5))continue;
                if(d.showGrid)AddLine(s,xx,top-7,xx,bottom,t.Grid,1,false,"grid");
                DateTime tick=origin.AddDays(days);
                AddText(s,Math.Max(left-22,Math.Min(right-72,xx-22)),230,94,28,tick.ToString(span>365?"MM/yyyy":"dd/MM/yy",Pt),12,t.Muted,false,"left","dateTick",-1);
            }
            for(int i=0;i<d.rows.Count;i++) {
                ChartRow row=d.rows[i];double cy=top+slot*(i+.5),h=Math.Min(23,slot*.55);
                double labelHeight=Math.Min(30,slot*.84);
                AddText(s,64,cy-labelHeight/2,260,labelHeight,row.label,d.rows.Count>16?11:15,t.Ink,true,"left","category",i);
                double xx=x(ProfessionalDate(row.startDate,false)),end=left+((ProfessionalDate(row.endDate,false)-origin).TotalDays+1)/span*(right-left),width=end-xx;
                string color=ProfessionalColor(d,t,i);
                AddRect(s,xx,cy-h/2,width,h,Mix(t.Background,color,.25),"ganttBar",i);
                AddRect(s,xx,cy-h/2,width*row.progress/100,h,color,"ganttProgress",i);
                if(d.showLabels)AddText(s,1148,cy-labelHeight/2,68,labelHeight,FormatNumber(row.progress,d.decimals)+"%",d.rows.Count>16?11:15,t.Ink,true,"right","value",i);
            }
            AddText(s,64,631,1152,22,"Período: "+origin.ToString("dd/MM/yyyy",Pt)+" a "+last.ToString("dd/MM/yyyy",Pt)+" · preenchimento indica o progresso informado.",14,t.Muted,false,"left","insight",-1);
        }
        static ExecutiveScale ProfessionalScale(IEnumerable<double> input) {
            // Normalizing before taking ticks avoids underflow for subnormal coordinates.
            List<double> values=input.ToList();double magnitude=values.Max(v=>Math.Abs(v));
            if(magnitude<1e-200)return new ExecutiveScale {Magnitude=1,Low=-1,High=1,Step=.5};
            return Scale(values);
        }
        static void ProfessionalPoints(Scene s,ChartDocument d,ExecutiveTheme t) {
            const double left=128,right=1178,top=270,bottom=565;
            ExecutiveScale xs=ProfessionalScale(d.rows.Select(r=>r.x)),ys=ProfessionalScale(d.rows.Select(r=>r.y));
            Func<double,double> x=value=>xs.Map(value,left,right),y=value=>ys.Map(value,bottom,top);
            VerticalAxis(s,d,t,ys,left,right,top,bottom);
            int ticks=(int)Math.Round((xs.High-xs.Low)/xs.Step);
            HashSet<string> labels=new HashSet<string>();
            for(int i=0;i<=ticks && i<20;i++) {
                double v=(xs.Low+i*xs.Step)*xs.Magnitude;if(Math.Abs(xs.Low+i*xs.Step)<xs.Step*.001)v=0;
                string text=ExecutiveValue(d,v,true);if(!labels.Add(text))continue;
                double xx=x(v);
                if(v==0 || d.showGrid)AddLine(s,xx,top,xx,bottom,v==0?t.Axis:t.Grid,v==0?1.3:1,false,v==0?"zeroAxis":"grid");
                AddText(s,xx-44,578,88,24,text,12,t.Muted,false,"center","xTick",-1);
            }
            AddText(s,64,232,300,24,"EIXO Y",12,t.Muted,true,"left","axisTitle",-1);
            AddText(s,1088,610,128,22,"EIXO X",12,t.Muted,true,"right","axisTitle",-1);
            double maxSize=d.type=="bubble"?d.rows.Max(r=>r.size):1;
            IEnumerable<int> order=Enumerable.Range(0,d.rows.Count);
            if(d.type=="bubble")order=order.OrderByDescending(i=>d.rows[i].size);
            foreach(int i in order) {
                ChartRow row=d.rows[i];double radius=d.type=="bubble"?29*Math.Sqrt(row.size)/Math.Sqrt(maxSize):5;
                string color=ProfessionalColor(d,t,i);
                AddEllipse(s,x(row.x)-radius,y(row.y)-radius,radius*2,radius*2,d.type=="bubble"?Mix(t.Background,color,.67):color,d.type=="bubble"?"bubble":"point",i);
            }
            if(d.showLabels)ProfessionalPointLabels(s,d,t,x,y);
            string insight=d.type=="bubble"?"Área proporcional ao tamanho · maior bolha: "+ExecutiveValue(d,maxSize,false)+".":"X e Y usam escalas independentes.";
            insight+=d.showLabels && d.rows.Count>20?" Rótulos numéricos seguem a ordem das linhas da tabela.":" Cada ponto representa uma categoria.";
            AddText(s,64,632,1152,22,insight,14,t.Muted,false,"left","insight",-1);
        }
        static void ProfessionalPointLabels(Scene s,ChartDocument d,ExecutiveTheme t,Func<double,double> x,Func<double,double> y) {
            if(d.rows.Count>20){ProfessionalNumberLabels(s,d,t,x,y);return;}
            // Reserve separate text frames; the point coordinates themselves never move.
            List<double[]> occupied=new List<double[]>();int n=d.rows.Count;
            double largest=d.type=="bubble"?d.rows.Max(r=>r.size):1;
            for(int i=0;i<n;i++) {
                ChartRow row=d.rows[i];double px=x(row.x),py=y(row.y),bestX=0,bestY=0,distance=Double.MaxValue;
                double w=152,h=36;
                for(int gy=0;gy<8;gy++)for(int gx=0;gx<8;gx++) {
                    double xx=112+gx*138,yy=256+gy*43;
                    if(occupied.Any(b=>xx<b[0]+w+2 && xx+w+2>b[0] && yy<b[1]+h+2 && yy+h+2>b[1]))continue;
                    double dd=(xx+w/2-px)*(xx+w/2-px)+(yy+h/2-py)*(yy+h/2-py);
                    foreach(ChartRow other in d.rows) {
                        double radius=d.type=="bubble"?29*Math.Sqrt(other.size)/Math.Sqrt(largest):5;
                        if(x(other.x)+radius+4>xx && x(other.x)-radius-4<xx+w && y(other.y)+radius+4>yy && y(other.y)-radius-4<yy+h)dd+=100000000;
                    }
                    if(dd<distance){distance=dd;bestX=xx;bestY=yy;}
                }
                occupied.Add(new[]{bestX,bestY});
                double edgeX=Math.Max(bestX,Math.Min(bestX+w,px)),edgeY=Math.Max(bestY,Math.Min(bestY+h,py));
                double radiusSelf=d.type=="bubble"?29*Math.Sqrt(row.size)/Math.Sqrt(largest):5;
                double length=Math.Sqrt((edgeX-px)*(edgeX-px)+(edgeY-py)*(edgeY-py));
                if(length>radiusSelf)AddLine(s,px+(edgeX-px)*radiusSelf/length,py+(edgeY-py)*radiusSelf/length,edgeX,edgeY,t.Axis,.7,false,"labelLeader");
                AddText(s,bestX,bestY,w,h,Wrap(row.label,20),13,t.Ink,true,"center","pointLabel",i);
            }
        }
        static void ProfessionalNumberLabels(Scene s,ChartDocument d,ExecutiveTheme t,Func<double,double> x,Func<double,double> y) {
            const double width=30,height=18;
            List<double[]> occupied=new List<double[]>();double largest=d.type=="bubble"?d.rows.Max(r=>r.size):1;
            for(int i=0;i<d.rows.Count;i++) {
                ChartRow row=d.rows[i];double px=x(row.x),py=y(row.y),radius=d.type=="bubble"?29*Math.Sqrt(row.size)/Math.Sqrt(largest):5;
                double bestX=px,bestY=py,bestScore=Double.MaxValue;
                for(int ring=0;ring<13;ring++)for(int angle=0;angle<24;angle++) {
                    double distance=radius+17+ring*16,a=angle*Math.PI/12;
                    double xx=Math.Max(102,Math.Min(1186,px+distance*Math.Cos(a)-width/2)),yy=Math.Max(252,Math.Min(604,py+distance*Math.Sin(a)-height/2));
                    if(occupied.Any(b=>xx<b[0]+width+2 && xx+width+2>b[0] && yy<b[1]+height+2 && yy+height+2>b[1]))continue;
                    double score=Math.Pow(xx+width/2-px,2)+Math.Pow(yy+height/2-py,2);
                    foreach(ChartRow other in d.rows) {
                        double rr=d.type=="bubble"?29*Math.Sqrt(other.size)/Math.Sqrt(largest):5;
                        if(x(other.x)+rr+2>xx && x(other.x)-rr-2<xx+width && y(other.y)+rr+2>yy && y(other.y)-rr-2<yy+height)score+=1000000;
                    }
                    if(score<bestScore){bestScore=score;bestX=xx;bestY=yy;}
                }
                occupied.Add(new[]{bestX,bestY});
                double edgeX=Math.Max(bestX,Math.Min(bestX+width,px)),edgeY=Math.Max(bestY,Math.Min(bestY+height,py));
                double length=Math.Sqrt(Math.Pow(edgeX-px,2)+Math.Pow(edgeY-py,2));
                if(length>radius+4)AddLine(s,px+(edgeX-px)*radius/length,py+(edgeY-py)*radius/length,edgeX,edgeY,t.Axis,.6,false,"labelLeader");
                AddRect(s,bestX,bestY,width,height,t.Background,"badgeBackground",i);
                AddText(s,bestX,bestY,width,height,(i+1).ToString(CultureInfo.InvariantCulture),10,t.Ink,true,"center","pointLabel",i);
            }
        }
    }
}
