using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Xml;
using System.Xml.Linq;
namespace MelloCharts
{
    public sealed class WorkbookSource { public string path,sheet,range; }
    public sealed class ChartDocument { public WorkbookSource excelLink; public string seriesActual="Realizado",seriesForecast="Projetado",seriesBaseline="Baseline",fontFamily="Arial"; public int version; public string type,title,subtitle,source,unit,palette; public int decimals; public bool showLabels,showBaseline,includeTotal; public string style="executive"; public bool compactNumbers=false,showGrid=true; public ChartColors customColors; public List<ChartRow> rows; }
    public sealed class ChartRow { public string label,kind,startDate,endDate; public double actual,forecast,baseline,value,progress,x,y,size; }
    public sealed class Scene { public double Width=1280,Height=720; public List<SceneElement> Elements=new List<SceneElement>(); }
    public sealed class SceneElement { public string Type,Text,Color,Align,Role,FontFamily; public double X,Y,W,H,X2,Y2,Size; public double[] Points; public bool Bold,Dash; public int Category=-1; }
    public static partial class ChartCore
    {
        const int MaxText = 1000000;
        static readonly CultureInfo Pt = CultureInfo.GetCultureInfo("pt-BR");
        static ArgumentException Bad(string message) { return new ArgumentException(message); }
        static bool IsType(string type) { return ChartCatalog.Contains(type); }
        public static int RowLimit(string type) { return type=="gantt"?24:type=="scatter" || type=="bubble"?80:12; }
        static void Number(double n) { if (Double.IsNaN(n) || Double.IsInfinity(n) || Math.Abs(n)>1000000000) throw Bad("Use um número finito de até 1 bilhão por célula."); }
        static void TextField(string value, string name, int limit, bool required) { if(value==null || value.Length>limit || value.Any(c=>Char.IsControl(c)) || (required && value.Trim().Length==0)) throw Bad("Revise " + name + " (até " + limit + " caracteres)."); }
        public static ChartDocument Example(string type)
        {
            if(!IsType(type)) throw Bad("Tipo de gráfico inválido.");
            if(IsProfessionalType(type))return ProfessionalExample(type);
            ChartDocument d=new ChartDocument {version=1,type=type,title="Composição do investimento",subtitle="Realizado e saldo projetado por frente de trabalho",source="Exemplo com dados fictícios",unit="R$ mil",palette="ocean",decimals=0,showLabels=true,showBaseline=true,includeTotal=true,rows=new List<ChartRow>()};
            if(ChartCatalog.Get(type).SingleValue) {
                d.title=type=="pareto"?"Onde concentrar as ações de melhoria":"Composição do investimento";
                d.subtitle=type=="pareto"?"Categorias que mais contribuem para o total e participação acumulada":"Participação por frente de trabalho, com valores e percentuais";
                d.showBaseline=false; d.includeTotal=false;
                string[] names={"Engenharia","Equipamentos","Infraestrutura","Implantação","Serviços","Outros"};
                double[] values={420,300,160,100,60,40};
                for(int i=0;i<names.Length;i++) d.rows.Add(new ChartRow {label=names[i],value=values[i]});
            } else if(type=="waterfall") {
                d.title="Do baseline à projeção final"; d.subtitle="Contribuição de cada ajuste para o EAC";
                d.rows.Add(new ChartRow{label="Baseline",value=1200,kind="base"}); d.rows.Add(new ChartRow{label="Escopo",value=100,kind="delta"}); d.rows.Add(new ChartRow{label="Eficiência",value=-60,kind="delta"}); d.rows.Add(new ChartRow{label="Preços",value=40,kind="delta"}); d.rows.Add(new ChartRow{label="EAC",value=0,kind="total"});
            } else if(type=="buildup") {
                d.rows.Add(new ChartRow{label="Engenharia",actual=320,forecast=180,baseline=480}); d.rows.Add(new ChartRow{label="Infraestrutura",actual=450,forecast=250,baseline=660}); d.rows.Add(new ChartRow{label="Equipamentos",actual=260,forecast=210,baseline=460}); d.rows.Add(new ChartRow{label="Implantação",actual=130,forecast=80,baseline=200});
            } else {
                d.title="Desempenho por período"; d.subtitle="Comparação entre realizado, projeção e baseline";
                d.rows.Add(new ChartRow{label="Janeiro",actual=180,forecast=210,baseline=200}); d.rows.Add(new ChartRow{label="Fevereiro",actual=240,forecast=250,baseline=230}); d.rows.Add(new ChartRow{label="Março",actual=280,forecast=270,baseline=260}); d.rows.Add(new ChartRow{label="Abril",actual=310,forecast=320,baseline=300});
            }
            if(type=="bars") {
                d.title="Desempenho por área"; d.subtitle="Realizado e projetado comparados com a referência aprovada";
                string[] names={"Operações","Engenharia","Tecnologia","Suprimentos"};
                for(int i=0;i<d.rows.Count;i++) d.rows[i].label=names[i];
            }
            if(type=="lines" || type=="area" || type=="combo") {
                d.title="Evolução dos resultados"; d.subtitle="Realizado e projetado na mesma escala, com baseline de referência";
                d.rows.Add(new ChartRow{label="Maio",actual=350,forecast=370,baseline=340});
                d.rows.Add(new ChartRow{label="Junho",actual=390,forecast=420,baseline=380});
            }
            if(type=="stacked100") {
                d.title="Avanço por frente de trabalho"; d.subtitle="Parcela realizada e saldo projetado restante em cada categoria";
                d.showBaseline=false; d.includeTotal=false;
            }
            return d;
        }
        public static ChartDocument Validate(ChartDocument d)
        {
            if(d==null || d.version!=1 || !IsType(d.type)) throw Bad("Formato de projeto incompatível.");
            TextField(d.title,"título",80,true); TextField(d.subtitle,"subtítulo",110,false); TextField(d.source,"fonte",100,false); TextField(d.unit,"unidade",16,false);
            if((d.palette!="ocean" && d.palette!="forest" && d.palette!="ink") || d.decimals<0 || d.decimals>2) throw Bad("Opções de apresentação inválidas.");
            if(d.style!="executive" && d.style!="minimal" && d.style!="dark") throw Bad("Estilo de apresentação inválido.");
            ColorTools.Validate(d.customColors);
            TextField(d.seriesActual,"nome da primeira série",24,true);TextField(d.seriesForecast,"nome da segunda série",24,true);TextField(d.seriesBaseline,"nome da referência",24,true);
            if(!new[]{"Arial","Calibri","Aptos","Segoe UI"}.Contains(d.fontFamily))throw Bad("Fonte não suportada.");
            if(d.excelLink!=null) { TextField(d.excelLink.path,"arquivo Excel",1024,true);TextField(d.excelLink.sheet,"planilha Excel",31,true);TextField(d.excelLink.range,"intervalo Excel",64,true); }
            if(d.rows==null || d.rows.Count<1 || d.rows.Count>RowLimit(d.type)) throw Bad("Use de 1 a "+RowLimit(d.type)+" linhas neste modelo.");
            for(int i=0;i<d.rows.Count;i++) {
                ChartRow r=d.rows[i]; if(r==null) throw Bad("Linha " + (i+1) + ": dados ausentes."); TextField(r.label,"categoria da linha " + (i+1),32,true);
                // All numeric fields must remain serializable even when not used by this chart type.
                Number(r.actual); Number(r.forecast); Number(r.baseline); Number(r.value);Number(r.x);Number(r.y);Number(r.size);Number(r.progress);
                if(IsProfessionalType(d.type))ValidateProfessionalRow(r,d.type);
                else if(d.type=="waterfall") { if((i==0 && r.kind!="base") || (i>0 && r.kind!="delta" && r.kind!="total")) throw Bad("A primeira etapa deve ser a única base; use variação ou total nas demais."); }
                else if(ChartCatalog.Get(d.type).SingleValue) { if(r.value<0) throw Bad("Pizza, rosca e Pareto usam valores não negativos."); }
                else if(ChartCatalog.Get(d.type).NonNegative && (r.actual<0 || r.forecast<0 || (ChartCatalog.Get(d.type).Baseline && r.baseline<0))) throw Bad("Este gráfico usa valores não negativos.");
            }
            if(ChartCatalog.Get(d.type).SingleValue && d.rows.Sum(r=>r.value)<=0) throw Bad("Informe pelo menos um valor positivo para calcular a participação.");
            return d;
        }
        static JavaScriptSerializer Serializer() { return new JavaScriptSerializer {MaxJsonLength=MaxText,RecursionLimit=32}; }
        public static ChartDocument Clone(ChartDocument doc) { return Deserialize(Serialize(doc)); }
        public static string Serialize(ChartDocument doc) { Validate(doc); return Serializer().Serialize(doc); }
        static object Field(IDictionary<string,object> obj,string key) { object v; if(!obj.TryGetValue(key,out v) || v==null) throw Bad("Campo obrigatório ausente: " + key + "."); return v; }
        static string StringField(IDictionary<string,object> obj,string key) { object v=Field(obj,key); if(!(v is string)) throw Bad("Texto inválido: " + key + "."); return (string)v; }
        static bool BoolField(IDictionary<string,object> obj,string key) { object v=Field(obj,key); if(!(v is bool)) throw Bad("Opção inválida: " + key + "."); return (bool)v; }
        static double NumericField(IDictionary<string,object> obj,string key) { object v=Field(obj,key); if(!(v is int) && !(v is long) && !(v is double) && !(v is decimal)) throw Bad("Número inválido: " + key + "."); double n=Convert.ToDouble(v,CultureInfo.InvariantCulture); Number(n); return n; }
        static int IntField(IDictionary<string,object> obj,string key) { double n=NumericField(obj,key); if(n!=Math.Truncate(n)) throw Bad("Inteiro inválido: " + key + "."); return (int)n; }
        // JavaScriptSerializer also accepts JavaScript object literals. Gate it with the
        // JSON grammar so malformed projects and duplicate keys are never normalized silently.
        sealed class StrictJson
        {
            readonly string input;int position;
            public StrictJson(string text) {input=text;}
            bool Digit(char c) {return c>='0' && c<='9';}
            void Space() {while(position<input.Length && (input[position]==' ' || input[position]=='\t' || input[position]=='\r' || input[position]=='\n'))position++;}
            bool Take(char c) {Space();if(position<input.Length && input[position]==c){position++;return true;}return false;}
            void Need(char c) {if(!Take(c))throw Bad("Sintaxe JSON inválida na posição " + position + ".");}
            string ReadString()
            {
                Need('"');StringBuilder text=new StringBuilder();
                while(position<input.Length) {char c=input[position++];if(c=='"')return text.ToString();if(c<' ')throw Bad("Controle inválido em texto JSON.");
                    if(c=='\\') {if(position>=input.Length)throw Bad("Escape JSON incompleto.");char escaped=input[position++];switch(escaped) {case '"':case '\\':case '/':c=escaped;break;case 'b':c='\b';break;case 'f':c='\f';break;case 'n':c='\n';break;case 'r':c='\r';break;case 't':c='\t';break;case 'u':int code;if(position+4>input.Length || !Int32.TryParse(input.Substring(position,4),NumberStyles.AllowHexSpecifier,CultureInfo.InvariantCulture,out code))throw Bad("Escape Unicode JSON inválido.");position+=4;c=(char)code;break;default:throw Bad("Escape JSON inválido.");}}
                    text.Append(c);
                }
                throw Bad("Texto JSON sem fechamento de aspas.");
            }
            void ReadNumber()
            {
                if(position<input.Length && input[position]=='-')position++;
                if(position>=input.Length || !Digit(input[position]))throw Bad("Valor JSON inválido.");
                if(input[position]=='0')position++;else while(position<input.Length && Digit(input[position]))position++;
                if(position<input.Length && input[position]=='.'){position++;int start=position;while(position<input.Length && Digit(input[position]))position++;if(start==position)throw Bad("Fração JSON inválida.");}
                if(position<input.Length && (input[position]=='e' || input[position]=='E')){position++;if(position<input.Length && (input[position]=='+' || input[position]=='-'))position++;int start=position;while(position<input.Length && Digit(input[position]))position++;if(start==position)throw Bad("Expoente JSON inválido.");}
            }
            void Value(int depth)
            {
                if(depth>32)throw Bad("Projeto JSON excede a profundidade permitida.");Space();if(position>=input.Length)throw Bad("Valor JSON ausente.");char c=input[position];
                if(c=='{'){position++;HashSet<string> keys=new HashSet<string>(StringComparer.Ordinal);if(Take('}'))return;do{string key=ReadString();if(!keys.Add(key))throw Bad("Campo JSON duplicado: "+key+".");Need(':');Value(depth+1);}while(Take(','));Need('}');}
                else if(c=='['){position++;if(Take(']'))return;do{Value(depth+1);}while(Take(','));Need(']');}
                else if(c=='"')ReadString();
                else if(c=='t' || c=='f' || c=='n') {string word=c=='t'?"true":c=='f'?"false":"null";if(position+word.Length>input.Length || String.CompareOrdinal(input,position,word,0,word.Length)!=0)throw Bad("Valor JSON inválido.");position+=word.Length;}
                else ReadNumber();
            }
            public void Check() {Value(0);Space();if(position!=input.Length)throw Bad("Conteúdo após o projeto JSON.");}
        }
        public static ChartDocument Deserialize(string json)
        {
            if(String.IsNullOrWhiteSpace(json) || json.Length>MaxText) throw Bad("O projeto deve conter até 1 MB de texto JSON.");
            try {
                new StrictJson(json).Check();
                IDictionary<string,object> o=Serializer().DeserializeObject(json) as IDictionary<string,object>; if(o==null) throw Bad("Objeto de projeto inválido.");
                ChartDocument d=new ChartDocument {version=IntField(o,"version"),type=StringField(o,"type"),title=StringField(o,"title"),subtitle=StringField(o,"subtitle"),source=StringField(o,"source"),unit=StringField(o,"unit"),palette=StringField(o,"palette"),decimals=IntField(o,"decimals"),showLabels=BoolField(o,"showLabels"),showBaseline=BoolField(o,"showBaseline"),includeTotal=BoolField(o,"includeTotal"),rows=new List<ChartRow>()};
                // Optional presentation fields keep saved v0.3 projects compatible.
                if(o.ContainsKey("style"))d.style=StringField(o,"style");if(o.ContainsKey("compactNumbers"))d.compactNumbers=BoolField(o,"compactNumbers");if(o.ContainsKey("showGrid"))d.showGrid=BoolField(o,"showGrid");
                if(o.ContainsKey("customColors"))d.customColors=ColorTools.Read(o["customColors"]);
                if(o.ContainsKey("seriesActual"))d.seriesActual=StringField(o,"seriesActual");if(o.ContainsKey("seriesForecast"))d.seriesForecast=StringField(o,"seriesForecast");if(o.ContainsKey("seriesBaseline"))d.seriesBaseline=StringField(o,"seriesBaseline");if(o.ContainsKey("fontFamily"))d.fontFamily=StringField(o,"fontFamily");
                if(o.ContainsKey("excelLink") && o["excelLink"]!=null) { var link=o["excelLink"] as IDictionary<string,object>;if(link==null)throw Bad("Vínculo Excel inválido.");d.excelLink=new WorkbookSource {path=StringField(link,"path"),sheet=StringField(link,"sheet"),range=StringField(link,"range")}; }
                IList rows=Field(o,"rows") as IList; if(rows==null || rows.Count<1 || rows.Count>RowLimit(d.type)) throw Bad("Use de 1 a "+RowLimit(d.type)+" linhas neste modelo.");
                foreach(object item in rows) { IDictionary<string,object> r=item as IDictionary<string,object>; if(r==null) throw Bad("Linha de projeto inválida."); ChartRow row=new ChartRow {label=StringField(r,"label")};
                    if(IsProfessionalType(d.type))ReadProfessionalRow(r,row,d.type);
                    else if(d.type=="waterfall") { row.kind=StringField(r,"kind"); row.value=NumericField(r,"value"); }
                    else if(ChartCatalog.Get(d.type).SingleValue) { row.value=NumericField(r,"value"); }
                    else { row.actual=NumericField(r,"actual"); row.forecast=NumericField(r,"forecast"); row.baseline=NumericField(r,"baseline"); }
                    // Validate optional numeric fields too, instead of silently discarding corrupt values.
                    if(r.ContainsKey("actual")) row.actual=NumericField(r,"actual"); if(r.ContainsKey("forecast")) row.forecast=NumericField(r,"forecast"); if(r.ContainsKey("baseline")) row.baseline=NumericField(r,"baseline"); if(r.ContainsKey("value")) row.value=NumericField(r,"value");
                    if(r.ContainsKey("x"))row.x=NumericField(r,"x");if(r.ContainsKey("y"))row.y=NumericField(r,"y");if(r.ContainsKey("size"))row.size=NumericField(r,"size");if(r.ContainsKey("progress"))row.progress=NumericField(r,"progress");
                    if(r.ContainsKey("kind") && r["kind"]!=null) row.kind=StringField(r,"kind"); d.rows.Add(row);
                }
                return Validate(d);
            } catch(ArgumentException) { throw; } catch(Exception e) { if(e is InvalidOperationException || e is FormatException || e is OverflowException) throw Bad("JSON inválido: " + e.Message); throw; }
        }
        public static double ParseNumber(string value)
        {
            if(String.IsNullOrWhiteSpace(value) || value.Length>100) throw Bad("Informe um valor numérico.");
            string s=Regex.Replace(value.Trim(),"^R\\$\\s*",""); s=Regex.Replace(s,"\\s","");
            if(s.StartsWith("(") && s.EndsWith(")")) s="-"+s.Substring(1,s.Length-2);
            string sign=""; if(s.StartsWith("-") || s.StartsWith("+")) {sign=s.Substring(0,1);s=s.Substring(1);}
            string exponent=""; int ep=s.IndexOfAny(new char[]{'e','E'}); if(ep>=0) { exponent=s.Substring(ep);s=s.Substring(0,ep);if(!Regex.IsMatch(exponent,"^[eE][+-]?\\d{1,3}$")) throw Bad("Número inválido."); }
            if(s.Contains(",") && s.Contains(".")) {
                if(s.LastIndexOf(',')>s.LastIndexOf('.')) { if(!Regex.IsMatch(s,"^\\d{1,3}(?:\\.\\d{3})*(?:,\\d+)?$")) throw Bad("Número inválido."); s=s.Replace(".","").Replace(',','.'); }
                else { if(!Regex.IsMatch(s,"^\\d{1,3}(?:,\\d{3})*(?:\\.\\d+)?$")) throw Bad("Número inválido."); s=s.Replace(",",""); }
            } else if(s.Contains(",")) { if(!Regex.IsMatch(s,"^\\d+,\\d+$")) throw Bad("Número inválido."); s=s.Replace(',','.'); }
            else if(exponent.Length==0 && Regex.IsMatch(s,"^\\d{1,3}(?:\\.\\d{3})+$")) s=s.Replace(".","");
            else if(!Regex.IsMatch(s,"^\\d+(?:\\.\\d+)?$")) throw Bad("Número inválido.");
            double n; if(!Double.TryParse(sign+s+exponent,NumberStyles.Float,CultureInfo.InvariantCulture,out n)) throw Bad("Número inválido."); Number(n); return n;
        }
        static string Norm(string value) { StringBuilder b=new StringBuilder(); foreach(char c in value.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD)) if(CharUnicodeInfo.GetUnicodeCategory(c)!=UnicodeCategory.NonSpacingMark) b.Append(c); return b.ToString(); }
        static string HeaderKey(string s) { switch(Norm(s)) {case "categoria":case "category":case "etapa":case "pep":case "projeto":case "label":return "label";case "realizado":case "actual":return "actual";case "projetado":case "forecast":return "forecast";case "baseline":return "baseline";case "valor":case "value":return "value";case "tipo":case "kind":return "kind";default:return Norm(s);} }
        static char Separator(string text)
        {
            bool quote=false; int tab=0,semi=0,comma=0;
            for(int i=0;i<text.Length;i++) {char c=text[i];if(c=='"') {if(quote && i+1<text.Length && text[i+1]=='"') i++;else quote=!quote;} else if(!quote) {if(c=='\r' || c=='\n') break;if(c=='\t')tab++;if(c==';')semi++;if(c==',')comma++;}}
            return tab>=semi && tab>=comma ? '\t' : semi>=comma ? ';' : ',';
        }
        static List<List<string>> ParseDelimited(string text)
        {
            if(text==null || text.Length>MaxText) throw Bad("A tabela deve ter até 1 MB de texto."); if(text.StartsWith("\uFEFF",StringComparison.Ordinal))text=text.Substring(1);
            char sep=Separator(text); List<List<string>> rows=new List<List<string>>(); List<string> row=new List<string>(); StringBuilder cell=new StringBuilder();bool quoted=false,closed=false;
            for(int i=0;i<text.Length;i++) {char c=text[i];
                if(quoted) {if(c=='"') {if(i+1<text.Length && text[i+1]=='"') {cell.Append('"');i++;}else{quoted=false;closed=true;}}else cell.Append(c);continue;}
                if(c=='"') {if(cell.Length>0 || closed)throw Bad("Aspas inválidas no arquivo.");quoted=true;}
                else if(c==sep) {row.Add(cell.ToString());cell.Length=0;closed=false;if(row.Count>64)throw Bad("A tabela tem colunas demais.");}
                else if(c=='\r' || c=='\n') {if(c=='\r' && i+1<text.Length && text[i+1]=='\n')i++;row.Add(cell.ToString());if(row.Any(v=>!String.IsNullOrWhiteSpace(v)))rows.Add(row);if(rows.Count>81)throw Bad("A tabela excede 80 linhas de dados.");row=new List<string>();cell.Length=0;closed=false;}
                else {if(closed && !Char.IsWhiteSpace(c))throw Bad("Conteúdo após fechamento de aspas.");if(!closed)cell.Append(c);}
            }
            if(quoted)throw Bad("Há uma célula com aspas não fechadas.");row.Add(cell.ToString());if(row.Any(v=>!String.IsNullOrWhiteSpace(v)))rows.Add(row);return rows;
        }
        static ChartDocument ImportRows(List<List<string>> rows,string type,bool invariantNumbers)
        {
            if(rows==null || rows.Count<2)throw Bad("Inclua cabeçalho e pelo menos uma linha.");
            if(!IsType(type))throw Bad("Tipo de gráfico inválido.");if(rows.Count>RowLimit(type)+1)throw Bad("Use até "+RowLimit(type)+" linhas neste modelo.");
            if(IsProfessionalType(type))return ImportProfessionalRows(rows,type,invariantNumbers); if(rows[0].Count>64)throw Bad("A tabela tem colunas demais.");
            string[] headers=rows[0].Select(HeaderKey).ToArray(); if(headers.Any(String.IsNullOrEmpty) || headers.Distinct().Count()!=headers.Length) throw Bad("Há cabeçalhos vazios ou duplicados.");
            string[] keys=type=="waterfall" ? new string[]{"label","value","kind"} : ChartCatalog.Get(type).SingleValue ? new string[]{"label","value"} : new string[]{"label","actual","forecast","baseline"};
            if(keys.Any(k=>Array.IndexOf(headers,k)<0))throw Bad("Cabeçalho esperado: "+ChartCatalog.Get(type).Header+".");
            ChartDocument d=Example(type);d.rows.Clear();
            for(int i=1;i<rows.Count;i++) {List<string> cells=rows[i];if(cells.Count!=headers.Length)throw Bad("Linha " + (i+1) + ": quantidade de colunas diferente do cabeçalho.");ChartRow r=new ChartRow();
                foreach(string key in keys) {string v=cells[Array.IndexOf(headers,key)];if(key=="label")r.label=Regex.Replace(v.Trim(),"\\s+"," ");else if(key=="kind") {switch(Norm(v)){case "base":r.kind="base";break;case "variacao":case "delta":r.kind="delta";break;case "subtotal":case "total":r.kind="total";break;default:throw Bad("Tipo de etapa inválido na linha " + (i+1) + ".");}}else {double n;if(invariantNumbers) {if(!Double.TryParse(v,NumberStyles.Float,CultureInfo.InvariantCulture,out n) || String.IsNullOrWhiteSpace(v))throw Bad("Número inválido na linha " + (i+1) + ".");Number(n);}else n=ParseNumber(v);switch(key){case "actual":r.actual=n;break;case "forecast":r.forecast=n;break;case "baseline":r.baseline=n;break;case "value":r.value=n;break;}}} d.rows.Add(r);
            }
            return Validate(d);
        }
        public static ChartDocument ImportTable(List<List<string>> rows,string type,bool invariantNumbers) { return ImportRows(rows,type,invariantNumbers); }
        public static ChartDocument ImportDelimited(string text,string type) { return ImportRows(ParseDelimited(text),type,false); }
        static string PartPath(string parent,string target)
        {
            if(String.IsNullOrWhiteSpace(target) || target.IndexOfAny(new char[]{':','\\','?','#','%'})>=0) throw new InvalidDataException("Relação XLSX inválida.");
            string combined=target.StartsWith("/")?target.Substring(1):parent.Substring(0,parent.LastIndexOf('/')+1)+target;
            List<string> parts=new List<string>();foreach(string part in combined.Split('/')) {if(part=="" || part==".")continue;if(part=="..") {if(parts.Count==0)throw new InvalidDataException("Relação fora do pacote XLSX.");parts.RemoveAt(parts.Count-1);}else parts.Add(part);}
            if(parts.Count==0)throw new InvalidDataException("Relação XLSX vazia.");return String.Join("/",parts.ToArray());
        }
        static XDocument ReadXml(ZipArchive zip,string path,bool required)
        {
            ZipArchiveEntry entry=zip.GetEntry(path);if(entry==null){if(required)throw new InvalidDataException("Parte XLSX ausente: " + path + ".");return null;}
            if(entry.Length>8000000)throw new InvalidDataException("Parte XLSX excede 8 MB descomprimidos.");
            XmlReaderSettings settings=new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=8000000,MaxCharactersFromEntities=0};
            using(Stream stream=entry.Open()) using(XmlReader reader=XmlReader.Create(stream,settings))return XDocument.Load(reader,LoadOptions.None);
        }
        static string RelationPath(string part) { int slash=part.LastIndexOf('/');return part.Substring(0,slash+1)+"_rels/"+part.Substring(slash+1)+".rels"; }
        static string Attr(XElement e,string name) {return (string)e.Attribute(name) ?? "";}
        static XNamespace SheetNamespace(XDocument document)
        {
            if(document.Root==null)throw new InvalidDataException("XML XLSX vazio.");string ns=document.Root.Name.NamespaceName;
            if(ns!="http://schemas.openxmlformats.org/spreadsheetml/2006/main" && ns!="http://purl.oclc.org/ooxml/spreadsheetml/main")throw new InvalidDataException("Namespace XLSX inválido.");return (XNamespace)ns;
        }
        static IEnumerable<XElement> Relationships(XDocument document)
        {
            if(document==null || document.Root==null || document.Root.Name!=XName.Get("Relationships","http://schemas.openxmlformats.org/package/2006/relationships"))throw new InvalidDataException("Relações XLSX inválidas.");
            List<XElement> rs=document.Root.Elements(document.Root.Name.Namespace+"Relationship").ToList();HashSet<string> ids=new HashSet<string>();foreach(XElement r in rs)if(Attr(r,"Id")=="" || !ids.Add(Attr(r,"Id")))throw new InvalidDataException("Relação XLSX duplicada ou vazia.");return rs;
        }
        static string ResolveRelation(string source,XElement relationship)
        {
            if(relationship==null || (Attr(relationship,"TargetMode")!="" && Attr(relationship,"TargetMode")!="Internal"))throw new InvalidDataException("Relações externas não são permitidas na importação XLSX.");return PartPath(source,Attr(relationship,"Target"));
        }
        static int Coordinate(string text,out int row)
        {
            Match match=Regex.Match(text,"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$");if(!match.Success)throw new InvalidDataException("Referência de célula XLSX inválida.");
            int col=0;foreach(char c in match.Groups[1].Value.ToUpperInvariant())col=col*26+c-'A'+1;row=Int32.Parse(match.Groups[2].Value,CultureInfo.InvariantCulture);
            if(col>64 || row>256)throw new InvalidDataException("A importação XLSX admite até 256 linhas e 64 colunas ocupadas.");return col;
        }
        static string RichText(XElement container,XNamespace ns) {if(container==null)return "";return String.Concat(container.Elements(ns+"t").Select(e=>e.Value).Concat(container.Elements(ns+"r").SelectMany(r=>r.Elements(ns+"t")).Select(e=>e.Value)));}
        static string CellValue(XElement cell,XNamespace ns,List<string> shared)
        {
            XElement value=cell.Element(ns+"v");string type=Attr(cell,"t");
            if(type=="e")throw new InvalidDataException("A planilha contém erro de célula. Corrija-o no Excel antes de importar.");
            if(cell.Element(ns+"f")!=null && (value==null || String.IsNullOrWhiteSpace(value.Value)))throw new InvalidDataException("A planilha contém fórmula sem valor em cache. Recalcule e salve no Excel antes de importar.");
            if(type=="inlineStr")return RichText(cell.Element(ns+"is"),ns);
            string text=value==null?"":value.Value;
            if(type=="s") {int index;if(!Int32.TryParse(text,NumberStyles.None,CultureInfo.InvariantCulture,out index) || index<0 || index>=shared.Count)throw new InvalidDataException("Índice de texto compartilhado inválido.");return shared[index];}
            if(type!="" && type!="n" && type!="str" && type!="b" && type!="d")throw new InvalidDataException("Tipo de célula XLSX incompatível.");return text;
        }
        public static ChartDocument ImportXlsx(string path,string type)
        {
            if(!IsType(type))throw Bad("Tipo de gráfico inválido.");if(String.IsNullOrWhiteSpace(path))throw Bad("Escolha um arquivo XLSX.");
            try {
                FileInfo info=new FileInfo(path);if(!info.Exists || info.Length>10000000)throw new InvalidDataException("Escolha um arquivo XLSX de até 10 MB.");
                using(FileStream file=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read)) using(ZipArchive zip=new ZipArchive(file,ZipArchiveMode.Read)) {
                    if(zip.Entries.Count>256)throw new InvalidDataException("O pacote XLSX tem partes demais.");long expanded=0;HashSet<string> names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach(ZipArchiveEntry entry in zip.Entries) {if(entry.Length>8000000 || entry.Length<0 || !names.Add(entry.FullName))throw new InvalidDataException("Parte XLSX duplicada ou grande demais.");expanded+=entry.Length;if(expanded>24000000)throw new InvalidDataException("O pacote XLSX excede 24 MB descomprimidos.");}
                    string workbookPath="xl/workbook.xml";XDocument rootRels=ReadXml(zip,"_rels/.rels",false);
                    if(rootRels!=null) {XElement rootRelation=Relationships(rootRels).FirstOrDefault(r=>Attr(r,"Type").EndsWith("/officeDocument",StringComparison.Ordinal));workbookPath=ResolveRelation("",rootRelation);}
                    XDocument workbook=ReadXml(zip,workbookPath,true);XNamespace ns=SheetNamespace(workbook);XElement sheets=workbook.Root.Element(ns+"sheets");XElement first=sheets==null?null:sheets.Elements(ns+"sheet").FirstOrDefault();if(first==null)throw new InvalidDataException("A pasta não contém planilhas.");
                    XAttribute id=first.Attributes().FirstOrDefault(a=>a.Name.LocalName=="id" && (a.Name.NamespaceName=="http://schemas.openxmlformats.org/officeDocument/2006/relationships" || a.Name.NamespaceName=="http://purl.oclc.org/ooxml/officeDocument/relationships"));if(id==null)throw new InvalidDataException("A primeira aba não possui relação válida.");
                    List<XElement> rels=Relationships(ReadXml(zip,RelationPath(workbookPath),true)).ToList();XElement sheetRelation=rels.FirstOrDefault(r=>Attr(r,"Id")==id.Value && Attr(r,"Type").EndsWith("/worksheet",StringComparison.Ordinal));string sheetPath=ResolveRelation(workbookPath,sheetRelation);
                    List<string> shared=new List<string>();XElement stringsRelation=rels.FirstOrDefault(r=>Attr(r,"Type").EndsWith("/sharedStrings",StringComparison.Ordinal));string sharedPath=stringsRelation==null?PartPath(workbookPath,"sharedStrings.xml"):ResolveRelation(workbookPath,stringsRelation);XDocument strings=ReadXml(zip,sharedPath,stringsRelation!=null);
                    if(strings!=null) {XNamespace sn=SheetNamespace(strings);foreach(XElement si in strings.Root.Elements(sn+"si")) {if(shared.Count>=16000)throw new InvalidDataException("A planilha tem textos compartilhados demais.");string text=RichText(si,sn);if(text.Length>100000)throw new InvalidDataException("Texto de célula grande demais.");shared.Add(text);}}
                    XDocument sheet=ReadXml(zip,sheetPath,true);XNamespace sheetNs=SheetNamespace(sheet);XElement data=sheet.Root.Element(sheetNs+"sheetData");if(data==null)throw new InvalidDataException("A primeira aba está vazia.");
                    SortedDictionary<int,Dictionary<int,string>> grid=new SortedDictionary<int,Dictionary<int,string>>();int previousRow=0,minCol=65,maxCol=0,cellCount=0;
                    foreach(XElement rowElement in data.Elements(sheetNs+"row")) {
                        int rowIndex=previousRow+1;string rowAttribute=Attr(rowElement,"r");if(rowAttribute!="" && !Int32.TryParse(rowAttribute,NumberStyles.None,CultureInfo.InvariantCulture,out rowIndex))throw new InvalidDataException("Linha XLSX inválida.");if(rowIndex<1 || rowIndex>256 || grid.ContainsKey(rowIndex))throw new InvalidDataException("A importação XLSX admite até 256 linhas sem duplicatas.");previousRow=rowIndex;
                        Dictionary<int,string> cells=new Dictionary<int,string>();int previousCol=0;
                        foreach(XElement cell in rowElement.Elements(sheetNs+"c")) {
                            if(++cellCount>16384)throw new InvalidDataException("A planilha contém células demais.");int col=previousCol+1;string reference=Attr(cell,"r");if(reference!="") {int cellRow;col=Coordinate(reference,out cellRow);if(cellRow!=rowIndex)throw new InvalidDataException("Referência de célula não corresponde à linha.");}if(col>64 || cells.ContainsKey(col))throw new InvalidDataException("Coluna XLSX fora do limite ou duplicada.");previousCol=col;
                            string value=CellValue(cell,sheetNs,shared);if(value.Length>100000)throw new InvalidDataException("Texto de célula grande demais.");cells.Add(col,value);if(!String.IsNullOrWhiteSpace(value)){minCol=Math.Min(minCol,col);maxCol=Math.Max(maxCol,col);}
                        }
                        grid.Add(rowIndex,cells);
                    }
                    List<List<string>> table=new List<List<string>>();foreach(Dictionary<int,string> cells in grid.Values) {if(!cells.Values.Any(v=>!String.IsNullOrWhiteSpace(v)))continue;List<string> row=new List<string>();for(int col=minCol;col<=maxCol;col++) {string value;row.Add(cells.TryGetValue(col,out value)?value:"");}table.Add(row);if(table.Count>13)throw Bad("Use de 1 a 12 categorias.");}
                    return ImportRows(table,type,true);
                }
            } catch(XmlException e) {throw new InvalidDataException("XML XLSX inválido: " + e.Message,e);} catch(NotSupportedException e) {throw new InvalidDataException("Formato XLSX incompatível.",e);}
        }
        static string CsvLabel(string label) { if(label.Length>0 && "=+@-".IndexOf(label.TrimStart()[0])>=0)label="'"+label;return "\""+label.Replace("\"","\"\"")+"\""; }
        static string CsvNumber(double value) { return value.ToString("R",CultureInfo.InvariantCulture).Replace('.',','); }
        public static string ToCsv(ChartDocument d)
        {
            Validate(d);if(IsProfessionalType(d.type))return ProfessionalCsv(d);StringBuilder b=new StringBuilder("\uFEFF"+ChartCatalog.Get(d.type).Header);
            // Totals are automatic: exporting the ignored zero placeholder also permits
            // accumulated totals larger than the per-input-cell limit to round-trip.
            foreach(ChartRow r in d.rows) {b.Append("\r\n").Append(CsvLabel(r.label));if(d.type=="waterfall")b.Append(';').Append(CsvNumber(r.kind=="total"?0:r.value)).Append(';').Append(r.kind=="delta"?"variacao":r.kind);else if(ChartCatalog.Get(d.type).SingleValue)b.Append(';').Append(CsvNumber(r.value));else b.Append(';').Append(CsvNumber(r.actual)).Append(';').Append(CsvNumber(r.forecast)).Append(';').Append(CsvNumber(r.baseline));}return b.ToString();
        }
        sealed class Palette
        {
            public string Actual,Forecast,Positive,Negative,Baseline;
            public Palette(string a,string f,string p,string n,string b) {Actual=a;Forecast=f;Positive=p;Negative=n;Baseline=b;}
        }
        static Palette Colors(string name,bool dark) {
            if(dark){if(name=="forest")return new Palette("#6CCDB0","#3D7D68","#E8B17B","#6CCDBB","#E8B17B");if(name=="ink")return new Palette("#B7C9E8","#62768F","#E8B17B","#81C5BC","#E8B17B");return new Palette("#69B8E8","#397D90","#E8B17B","#6CCDBB","#E8B17B");}
            if(name=="forest")return new Palette("#286153","#B4D6C6","#BD8050","#66A899","#A86D40");if(name=="ink")return new Palette("#3C4A63","#C2CBDC","#A97869","#839FA2","#AE7759");return new Palette("#244F73","#8ACBBF","#BD8050","#66A899","#B2763E");
        }
        sealed class Bar
        {
            public string Label,Kind; public double Actual,Forecast,Baseline,From,To,Value; public bool Total;
        }
        static List<Bar> Analyze(ChartDocument d)
        {
            List<Bar> bars=new List<Bar>();double running=0,actual=0,forecast=0,baseline=0;
            foreach(ChartRow r in d.rows) {Bar b=new Bar {Label=r.label,Kind=r.kind,Actual=r.actual,Forecast=r.forecast,Baseline=r.baseline};
                if(d.type=="waterfall") {if(r.kind=="base")running=r.value;else if(r.kind=="delta"){b.From=running;running+=r.value;}b.To=running;b.Value=r.kind=="delta"?r.value:running;}
                else {b.To=r.actual+r.forecast;actual+=r.actual;forecast+=r.forecast;baseline+=r.baseline;}bars.Add(b);
            }
            if(d.type=="buildup" && d.includeTotal)bars.Add(new Bar{Label="Total",Actual=actual,Forecast=forecast,Baseline=baseline,To=actual+forecast,Total=true});return bars;
        }
        static string Wrap(string value,int chars)
        {
            List<string> lines=new List<string>();string line="";foreach(string word in Regex.Split(value.Trim(),"\\s+")){if(word.Length>chars){if(line!=""){lines.Add(line);line="";}for(int i=0;i<word.Length;i+=chars)lines.Add(word.Substring(i,Math.Min(chars,word.Length-i)));}else if((line+" "+word).Trim().Length>chars){lines.Add(line);line=word;}else line=(line+" "+word).Trim();}if(line!="")lines.Add(line);return String.Join("\n",lines.ToArray());
        }
        public static string FormatNumber(double number,int decimals) {return number.ToString("N"+Math.Max(0,Math.Min(6,decimals)),Pt);}
        static string ValueLabel(double value,int decimals,bool compact) {
            if(!compact || Math.Abs(value)<1000)return FormatNumber(value,decimals);
            double divisor=Math.Abs(value)>=1000000000?1000000000:Math.Abs(value)>=1000000?1000000:1000;int digits=Math.Max(1,decimals);
            if(Math.Abs(Math.Round(value/divisor,digits))>=1000 && divisor<1000000000)divisor*=1000;
            return FormatNumber(value/divisor,digits)+(divisor==1000?" mil":divisor==1000000?" mi":" bi");
        }
        static double FitText(string text,double width,double height,double preferred,bool bold)
        {
            // Conservative Arial advances and line spacing keep the model independent
            // from GDI/Office while fitting the same fixed frames in both renderers.
            string[] lines=text.Split('\n');double maximum=0;
            foreach(string line in lines) {double units=0;foreach(char c in line) {if(c==' ')units+=0.34;else if("ilI.,:;!'|".IndexOf(c)>=0)units+=0.3;else if("WMmw@%".IndexOf(c)>=0 || c>0x2fff)units+=1.02;else if(Char.IsUpper(c))units+=0.8;else units+=0.62;}maximum=Math.Max(maximum,units);}
            double byWidth=width/(maximum*(bold?1.17:1.12)+1.4);double byHeight=height/(Math.Max(1,lines.Length)*1.36);
            return Math.Min(preferred,Math.Min(byWidth,byHeight));
        }
        static void AddText(Scene scene,double x,double y,double w,double h,string value,double size,string color,bool bold,string align,string role,int category)
        {
            scene.Elements.Add(new SceneElement{Type="text",X=x,Y=y,W=w,H=h,Text=value,Size=FitText(value,w,h,size,bold),Color=color,Bold=bold,Align=align,Role=role,Category=category});
        }
        static void AddRect(Scene scene,double x,double y,double w,double h,string color,string role,int category)
        {
            if(h>0 && w>0)scene.Elements.Add(new SceneElement{Type="rect",X=x,Y=y,W=w,H=h,Color=color,Role=role,Category=category,Size=1});
        }
        static void AddLine(Scene scene,double x,double y,double x2,double y2,string color,double size,bool dash,string role)
        {
            scene.Elements.Add(new SceneElement{Type="line",X=x,Y=y,X2=x2,Y2=y2,W=Math.Abs(x2-x),H=Math.Abs(y2-y),Color=color,Size=size,Dash=dash,Role=role});
        }
        public static Scene BuildScene(ChartDocument d)
        {
            Validate(d);Scene s=IsProfessionalType(d.type)?BuildProfessionalScene(d):BuildStandardScene(d);
            foreach(SceneElement e in s.Elements)if(e.Type=="text") {
                e.FontFamily=d.fontFamily;
                if(e.Role=="legend" || e.Role=="seriesLabel") {
                    string[] old={"Projetado restante","Realizado","Projetado","Baseline"};string[] names={d.seriesForecast=="Projetado"?"Projetado restante":d.seriesForecast,d.seriesActual,d.seriesForecast,d.seriesBaseline};
                    for(int i=0;i<old.Length;i++)if(e.Text==old[i] || e.Text.StartsWith(old[i]+"\n",StringComparison.Ordinal)) {e.Text=names[i]+e.Text.Substring(old[i].Length);e.Size=FitText(e.Text,e.W,e.H,e.Size,e.Bold);break;}
                }
            }
            return s;
        }
        static Scene BuildStandardScene(ChartDocument d)
        {
            if(d!=null && d.type!="buildup" && d.type!="waterfall" && d.type!="columns") return BuildAdvancedScene(Validate(d));
            Validate(d);Scene s=new Scene();bool dark=d.style=="dark",minimal=d.style=="minimal";Palette p=Colors(d);List<Bar> bars=Analyze(d);SurfaceColors surface=new SurfaceColors(d);
            string background=surface.Background,ink=surface.Ink,muted=surface.Muted,grid=surface.Grid,axis=surface.Axis,band=surface.Band;
            AddRect(s,0,0,s.Width,s.Height,background,"background",-1);
            string chartName=d.type=="buildup"?"BUILD-UP":d.type=="waterfall"?"CASCATA":"COMPARATIVO";
            if(!minimal){AddRect(s,64,29,28,4,p.Actual,"accent",-1);AddText(s,104,21,1000,24,chartName,13,ColorText(d,p.Actual,background),true,"left","eyebrow",-1);}
            AddText(s,64,54,1152,78,Wrap(d.title,59),36,ink,true,"left","title",-1);
            AddText(s,64,137,1152,34,Wrap(d.subtitle,100),18,muted,false,"left","subtitle",-1);
            AddText(s,64,193,400,26,d.unit+(d.compactNumbers?" (abreviado)":""),15,muted,false,"left","unit",-1);
            string[] legendLabels=d.type=="waterfall"?new[]{"Base / total","Aumento","Redução"}:d.showBaseline?new[]{"Realizado",d.type=="buildup"?"Projetado restante":"Projetado","Baseline"}:new[]{"Realizado",d.type=="buildup"?"Projetado restante":"Projetado"};
            string[] legendColors=d.type=="waterfall"?new[]{p.Actual,p.Positive,p.Negative}:new[]{p.Actual,p.Forecast,p.Baseline};double legendX=610;
            for(int i=0;i<legendLabels.Length;i++){
                if(legendLabels[i]=="Baseline")AddLine(s,legendX,204,legendX+22,204,legendColors[i],2.5,true,"legend");else AddRect(s,legendX,198,13,13,legendColors[i],"legend",-1);
                AddText(s,legendX+30,191,195,28,legendLabels[i],16,muted,false,"left","legend",-1);legendX+=legendLabels[i]=="Projetado restante"?230:170;
            }
            const double left=112,right=1216,top=252,bottom=552;
            List<double> values=new List<double>{0};
            foreach(Bar b in bars){if(d.type=="waterfall"){values.Add(b.From);values.Add(b.To);}else if(d.type=="buildup"){values.Add(b.To);if(d.showBaseline)values.Add(b.Baseline);}else{values.Add(b.Actual);values.Add(b.Forecast);if(d.showBaseline)values.Add(b.Baseline);}}
            // Normalize before tick selection, including subnormal inputs and summed totals.
            double magnitude=values.Max(v=>Math.Abs(v));if(magnitude==0)magnitude=1;double lo=values.Min()/magnitude,hi=values.Max()/magnitude;if(lo==hi)hi=lo+1;
            double span=hi-lo;if(lo<0)lo-=span*0.18;if(hi>0)hi+=span*0.19;double initialStep=(hi-lo)/5;double power=Math.Pow(10,Math.Floor(Math.Log10(initialStep)));double step=10*power;foreach(double n in new double[]{1,2,2.5,5,10})if(n*power>=initialStep){step=n*power;break;}
            // Prefer human-readable absolute intervals (500, 1,000...) while retaining
            // normalized fallback for subnormal doubles that cannot represent a nice step.
            double rawStep=initialStep*magnitude,rawPower=Math.Pow(10,Math.Floor(Math.Log10(rawStep)));
            if(rawStep>0 && rawPower>0){foreach(double n in new double[]{1,2,2.5,5,10})if(n*rawPower>=rawStep){double candidate=n*rawPower/magnitude;if(candidate>0 && !Double.IsInfinity(candidate))step=candidate;break;}}
            lo=Math.Floor(lo/step)*step;hi=Math.Ceiling(hi/step)*step;double lower=lo,upper=hi,mag=magnitude;
            Func<double,double> y=delegate(double value){return bottom-(value/mag-lower)/(upper-lower)*(bottom-top);};
            Func<double,string> label=delegate(double value){return ValueLabel(value,d.decimals,d.compactNumbers);};
            double slot=(right-left)/bars.Count,bw=Math.Min(116,slot*0.64);
            // Background emphasis precedes the axes, bars and labels, so it never hides them.
            for(int i=0;i<bars.Count;i++)if(bars[i].Total){double cx=left+slot*(i+0.5);if(!minimal)AddRect(s,cx-slot*0.46,top-14,slot*0.92,398,band,"totalBackground",i);AddLine(s,cx-slot*0.4,top-14,cx+slot*0.4,top-14,p.Actual,3,false,"totalAccent");}
            double actualStep=step*magnitude;int tickDecimals=actualStep>0 && actualStep<1?Math.Min(6,(int)Math.Ceiling(-Math.Log10(actualStep))+1):d.decimals;
            for(int places=0;places<=6;places++)if(Math.Abs(actualStep-Math.Round(actualStep,places))<=Math.Abs(actualStep)*1e-10){tickDecimals=Math.Max(d.decimals,places);break;}
            int tickCount=(int)Math.Round((hi-lo)/step);
            for(int i=0;i<=tickCount && i<20;i++){
                double tick=lo+i*step,v=Math.Abs(tick)<step*0.001?0:tick*magnitude,yy=bottom-(tick-lo)/(hi-lo)*(bottom-top);
                if(v==0 || d.showGrid)AddLine(s,left,yy,right,yy,v==0?axis:grid,v==0?1.2:1,false,v==0?"zeroAxis":"grid");
                AddText(s,12,yy-12,86,25,ValueLabel(v,tickDecimals,d.compactNumbers),15,muted,false,"right","tick",-1);
            }
            // The connected baseline sits behind grouped columns and their labels.
            if(d.type=="columns" && d.showBaseline)for(int i=0;i<bars.Count;i++){double cx=left+slot*(i+0.5),yy=y(bars[i].Baseline);if(i>0)AddLine(s,cx-slot,y(bars[i-1].Baseline),cx,yy,p.Baseline,2.5,true,"baseline");AddRect(s,cx-3,yy-3,6,6,p.Baseline,"baseline",i);}
            for(int i=0;i<bars.Count;i++){
                Bar b=bars[i];double cx=left+slot*(i+0.5),x=cx-bw/2;
                if(d.type=="waterfall"){
                    string color=b.Kind=="delta"?(b.Value>=0?p.Positive:p.Negative):p.Actual;
                    AddRect(s,x,Math.Min(y(b.From),y(b.To)),bw,Math.Abs(y(b.To)-y(b.From)),color,"bar",i);
                    if(i<bars.Count-1)AddLine(s,cx+bw/2,y(b.To),cx+slot-bw/2,y(b.To),axis,1,true,"connector");
                    if(d.showLabels){string sign=b.Kind=="delta" && b.Value>0?"+":"";AddText(s,cx-slot*0.46,Math.Min(y(b.From),y(b.To))-32,slot*0.92,27,sign+label(b.Value),bars.Count>8?16:21,ink,true,"center","value",i);}
                }else if(d.type=="buildup"){
                    AddRect(s,x,y(b.Actual),bw,y(0)-y(b.Actual),p.Actual,"bar",i);AddRect(s,x,y(b.To),bw,y(b.Actual)-y(b.To),p.Forecast,"bar",i);
                    if(d.showBaseline)AddLine(s,x-9,y(b.Baseline),x+bw+9,y(b.Baseline),p.Baseline,2.5,true,"baseline");
                    if(d.showLabels){
                        AddText(s,cx-slot*0.46,y(b.To)-33,slot*0.92,28,label(b.To),b.Total?24:bars.Count>8?16:21,ink,true,"center","value",i);
                        double actualH=y(0)-y(b.Actual),forecastH=y(b.Actual)-y(b.To);string av=label(b.Actual),fv=label(b.Forecast);
                        if(actualH>38 && FitText(av,bw-8,25,17,true)>=12)AddText(s,x+4,y(b.Actual)+actualH/2-12,bw-8,25,av,17,FillInk(d,p.Actual,dark?"#101E2E":"#FFFFFF"),true,"center","actualValue",i);
                        if(forecastH>38 && FitText(fv,bw-8,25,17,true)>=12)AddText(s,x+4,y(b.To)+forecastH/2-12,bw-8,25,fv,17,FillInk(d,p.Forecast,dark?"#FFFFFF":"#173E45"),true,"center","forecastValue",i);
                    }
                }else{
                    double w=bw*0.44;double[] series=new[]{b.Actual,b.Forecast};string[] colors=new[]{p.Actual,p.Forecast};
                    for(int j=0;j<2;j++){double pos=x+j*bw*0.56,v=series[j];AddRect(s,pos,Math.Min(y(v),y(0)),w,Math.Abs(y(0)-y(v)),colors[j],"bar",i);
                        // Separate half-slot frames prevent equal-height series labels from overlapping.
                        if(d.showLabels)AddText(s,pos+w/2-bw*0.26,v>=0?y(v)-29:y(v)+4,bw*0.52,25,label(v),bars.Count>7?14:18,ink,true,"center","value",i);
                    }
                }
                AddText(s,cx-slot*0.46,574,slot*0.92,66,Wrap(b.Label,Math.Max(6,(int)Math.Floor(slot/10))),bars.Count>8?15:18,b.Total?ink:muted,b.Total,"center","category",i);
            }
            AddLine(s,64,664,1216,664,grid,1,false,"footer");AddText(s,64,679,970,25,d.source,14,muted,false,"left","source",-1);AddText(s,1080,679,136,25,"Mello Charts",13,muted,false,"right","brand",-1);return s;
        }
    }
}
