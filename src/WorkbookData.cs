using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace MelloCharts
{
    // Reads only the bytes saved in a local XLSX. No Excel automation, recalculation or polling.
    public static class WorkbookData
    {
        const int MaxCells=16384;
        static InvalidDataException Bad(string text){return new InvalidDataException(text);}
        static string Attr(XElement e,string name){return (string)e.Attribute(name)??"";}
        static string LocalPath(string path)
        {
            if(String.IsNullOrWhiteSpace(path))throw new ArgumentException("Escolha um arquivo XLSX local salvo.");
            string full=Path.GetFullPath(path);
            if(full.StartsWith("\\\\",StringComparison.Ordinal) || full.StartsWith("//",StringComparison.Ordinal) || !String.Equals(Path.GetExtension(full),".xlsx",StringComparison.OrdinalIgnoreCase))throw Bad("Escolha um arquivo XLSX em disco local; caminhos de rede não são permitidos.");
            DriveInfo drive=new DriveInfo(Path.GetPathRoot(full));if(drive.DriveType==DriveType.Network)throw Bad("Copie o XLSX para um disco local antes de importar.");
            return full;
        }
        static string PartPath(string parent,string target)
        {
            if(String.IsNullOrWhiteSpace(target) || target.IndexOfAny(new[]{':','\\','?','#','%'})>=0 || target.Any(Char.IsControl))throw Bad("Relação XLSX inválida.");
            string combined=target.StartsWith("/",StringComparison.Ordinal)?target.Substring(1):parent.Substring(0,parent.LastIndexOf('/')+1)+target;
            List<string> pieces=new List<string>();foreach(string piece in combined.Split('/')){if(piece=="" || piece==".")continue;if(piece==".."){if(pieces.Count==0)throw Bad("Relação fora do pacote XLSX.");pieces.RemoveAt(pieces.Count-1);}else pieces.Add(piece);}
            if(pieces.Count==0)throw Bad("Relação XLSX vazia.");return String.Join("/",pieces.ToArray());
        }
        static string RelationPath(string part){int slash=part.LastIndexOf('/');return part.Substring(0,slash+1)+"_rels/"+part.Substring(slash+1)+".rels";}
        static XNamespace SheetNamespace(XDocument doc)
        {
            if(doc.Root==null)throw Bad("XML XLSX vazio.");string ns=doc.Root.Name.NamespaceName;
            if(ns!="http://schemas.openxmlformats.org/spreadsheetml/2006/main" && ns!="http://purl.oclc.org/ooxml/spreadsheetml/main")throw Bad("Namespace XLSX inválido.");return ns;
        }
        static List<XElement> Relationships(XDocument doc)
        {
            XNamespace ns="http://schemas.openxmlformats.org/package/2006/relationships";
            if(doc==null || doc.Root==null || doc.Root.Name!=ns+"Relationships")throw Bad("Relações XLSX inválidas.");
            List<XElement> result=doc.Root.Elements(ns+"Relationship").ToList();HashSet<string> ids=new HashSet<string>(StringComparer.Ordinal);
            foreach(XElement r in result){if(Attr(r,"Id")=="" || !ids.Add(Attr(r,"Id")))throw Bad("Relação XLSX duplicada ou vazia.");if(Attr(r,"TargetMode")!="" && Attr(r,"TargetMode")!="Internal")throw Bad("Relações externas não são permitidas na importação XLSX.");}
            return result;
        }
        sealed class Package:IDisposable
        {
            FileStream file;ZipArchive zip;
            public readonly string Path,WorkbookPath;public readonly List<XElement> Sheets,Relations;public readonly bool Date1904;
            public Package(string path)
            {
                try{
                    Path=LocalPath(path);file=new FileStream(Path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
                    if(file.Length>10000000)throw Bad("Escolha um arquivo XLSX de até 10 MB.");
                    zip=new ZipArchive(file,ZipArchiveMode.Read);
                    if(zip.Entries.Count>256)throw Bad("O pacote XLSX tem partes demais.");long expanded=0;HashSet<string> names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach(ZipArchiveEntry entry in zip.Entries){if(entry.Length<0 || entry.Length>8000000 || !names.Add(entry.FullName))throw Bad("Parte XLSX duplicada ou grande demais.");expanded+=entry.Length;if(expanded>24000000)throw Bad("O pacote XLSX excede 24 MB descomprimidos.");}
                    // Check all relationship files, including hyperlinks on unselected sheets.
                    foreach(ZipArchiveEntry entry in zip.Entries)if(entry.FullName.EndsWith(".rels",StringComparison.OrdinalIgnoreCase)){
                        int folder=entry.FullName.LastIndexOf("_rels/",StringComparison.Ordinal);
                        if(folder<0)throw Bad("Caminho de relações XLSX inválido.");string source=entry.FullName.Substring(0,folder)+entry.FullName.Substring(folder+6);source=source.Substring(0,source.Length-5);
                        foreach(XElement r in Relationships(ReadXml(entry.FullName,true)))PartPath(source,Attr(r,"Target"));
                    }
                    WorkbookPath="xl/workbook.xml";XDocument root=ReadXml("_rels/.rels",false);
                    if(root!=null){XElement relation=Relationships(root).FirstOrDefault(r=>Attr(r,"Type").EndsWith("/officeDocument",StringComparison.Ordinal));if(relation==null)throw Bad("Relação da pasta XLSX ausente.");WorkbookPath=PartPath("",Attr(relation,"Target"));}
                    XDocument workbook=ReadXml(WorkbookPath,true);XNamespace ns=SheetNamespace(workbook);XElement properties=workbook.Root.Element(ns+"workbookPr");string date1904=properties==null?"":Attr(properties,"date1904");Date1904=date1904=="1" || date1904=="true";
                    if(date1904!="" && date1904!="0" && date1904!="false" && !Date1904)throw Bad("Sistema de datas XLSX inválido.");XElement sheets=workbook.Root.Element(ns+"sheets");
                    Sheets=sheets==null?new List<XElement>():sheets.Elements(ns+"sheet").ToList();if(Sheets.Count==0)throw Bad("A pasta não contém planilhas.");
                    HashSet<string> sheetNames=new HashSet<string>(StringComparer.OrdinalIgnoreCase);foreach(XElement sheet in Sheets)if(String.IsNullOrWhiteSpace(Attr(sheet,"name")) || !sheetNames.Add(Attr(sheet,"name")))throw Bad("Nome de planilha vazio ou duplicado.");
                    Relations=Relationships(ReadXml(RelationPath(WorkbookPath),true));
                }catch(XmlException e){Dispose();throw new InvalidDataException("XML XLSX inválido: "+e.Message,e);}catch{Dispose();throw;}
            }
            public XDocument ReadXml(string path,bool required)
            {
                ZipArchiveEntry entry=zip.GetEntry(path);if(entry==null){if(required)throw Bad("Parte XLSX ausente: "+path+".");return null;}
                XmlReaderSettings settings=new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=8000000,MaxCharactersFromEntities=0};
                using(Stream stream=entry.Open())using(XmlReader reader=XmlReader.Create(stream,settings))return XDocument.Load(reader,LoadOptions.None);
            }
            public void Dispose(){if(zip!=null){zip.Dispose();zip=null;}if(file!=null){file.Dispose();file=null;}}
        }
        sealed class Area
        {
            public int Top,Left,Bottom,Right;
            public string Reference{get{return ColumnName(Left)+Top.ToString(CultureInfo.InvariantCulture)+":"+ColumnName(Right)+Bottom.ToString(CultureInfo.InvariantCulture);}}
            public bool Contains(int row,int col){return row>=Top && row<=Bottom && col>=Left && col<=Right;}
        }
        static string ColumnName(int col){string name="";while(col>0){col--;name=(char)('A'+col%26)+name;col/=26;}return name;}
        static int Coordinate(string text,out int row,bool allowDollar)
        {
            Match m=Regex.Match(text,allowDollar?"^\\$?([A-Za-z]{1,3})\\$?([1-9][0-9]{0,6})$":"^([A-Za-z]{1,3})([1-9][0-9]{0,6})$");
            if(!m.Success)throw Bad("Referência de célula XLSX inválida. Use um intervalo A1, como H500:K510.");
            int col=0;foreach(char c in m.Groups[1].Value.ToUpperInvariant())col=col*26+c-'A'+1;row=Int32.Parse(m.Groups[2].Value,CultureInfo.InvariantCulture);
            if(col>16384 || row>1048576)throw Bad("Referência fora dos limites da planilha Excel.");return col;
        }
        static void Bound(Area area,string type)
        {
            if(area.Left>area.Right || area.Top>area.Bottom)throw Bad("O início do intervalo deve preceder o fim.");
            if(area.Right-area.Left+1>64)throw Bad("Selecione no máximo 64 colunas, incluindo os cabeçalhos.");
            int limit=ChartCore.RowLimit(type);if(area.Bottom-area.Top+1>limit+1)throw Bad("Selecione cabeçalho e até "+limit+" linhas de dados. Informe um intervalo menor.");
        }
        static Area ParseArea(string range,string type)
        {
            if(String.IsNullOrWhiteSpace(range))return null;string[] split=range.Trim().Split(':');if(split.Length!=2)throw Bad("Use um intervalo A1 com início e fim, incluindo o cabeçalho: H500:K510.");
            Area area=new Area();area.Left=Coordinate(split[0],out area.Top,true);area.Right=Coordinate(split[1],out area.Bottom,true);Bound(area,type);return area;
        }
        static string RichText(XElement container,XNamespace ns){return container==null?"":String.Concat(container.Elements().Where(e=>e.Name==ns+"t" || e.Name==ns+"r").SelectMany(e=>e.Name==ns+"t"?new[]{e}:e.Elements(ns+"t")).Select(e=>e.Value));}
        static string Header(string text){return String.Concat(text.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD).Where(c=>CharUnicodeInfo.GetUnicodeCategory(c)!=UnicodeCategory.NonSpacingMark));}
        static bool DateHeader(string text){string key=Header(text);return key=="inicio" || key=="data inicial" || key=="startdate" || key=="fim" || key=="data final" || key=="enddate";}
        static bool PercentFormat(string format)
        {
            bool quoted=false;for(int i=0;i<format.Length;i++){char c=format[i];if(c=='\\' || c=='_' || c=='*'){i++;continue;}if(c=='"'){quoted=!quoted;continue;}if(quoted)continue;if(c=='['){int close=format.IndexOf(']',i+1);if(close>=0){i=close;continue;}}if(c=='%')return true;}return false;
        }
        sealed class CellStyle {public int Id;public string Format;public bool Percent;}
        static List<CellStyle> ReadStyles(Package package)
        {
            List<CellStyle> result=new List<CellStyle>();XElement relation=package.Relations.FirstOrDefault(r=>Attr(r,"Type").EndsWith("/styles",StringComparison.Ordinal));
            XDocument styles=package.ReadXml(relation==null?PartPath(package.WorkbookPath,"styles.xml"):PartPath(package.WorkbookPath,Attr(relation,"Target")),relation!=null);if(styles==null){result.Add(new CellStyle{Id=0,Format="General"});return result;}
            XNamespace ns=SheetNamespace(styles);Dictionary<int,string> formats=new Dictionary<int,string>();XElement numFmts=styles.Root.Element(ns+"numFmts");
            if(numFmts!=null)foreach(XElement format in numFmts.Elements(ns+"numFmt")){int id;if(!Int32.TryParse(Attr(format,"numFmtId"),NumberStyles.None,CultureInfo.InvariantCulture,out id) || formats.ContainsKey(id))throw Bad("Formato numérico XLSX inválido ou duplicado.");formats.Add(id,Attr(format,"formatCode"));}
            XElement xfs=styles.Root.Element(ns+"cellXfs");if(xfs!=null)foreach(XElement xf in xfs.Elements(ns+"xf")){if(result.Count>=65536)throw Bad("A planilha contém estilos demais.");int id;string raw=Attr(xf,"numFmtId");if(raw=="")id=0;else if(!Int32.TryParse(raw,NumberStyles.None,CultureInfo.InvariantCulture,out id))throw Bad("Índice de formato numérico XLSX inválido.");string format;if(!formats.TryGetValue(id,out format))format=id==0?"General":id==1?"0":id==49?"@":null;result.Add(new CellStyle{Id=id,Format=format,Percent=id==9 || id==10 || (format!=null && PercentFormat(format))});}
            if(result.Count==0)result.Add(new CellStyle{Id=0,Format="General"});return result;
        }
        static bool CategoryHeader(string text){string key=Header(text);return key=="categoria" || key=="category" || key=="atividade" || key=="etapa" || key=="pep" || key=="projeto" || key=="label";}
        static string CategoryValue(string value,CellStyle style,bool date1904)
        {
            if(String.IsNullOrWhiteSpace(value))return value;double number;
            if(!Double.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out number) || Double.IsNaN(number) || Double.IsInfinity(number))throw Bad("Categoria numérica XLSX inválida.");
            string format=style.Format;
            if(format=="General" || format=="@")return number.ToString("G15",CultureInfo.CurrentCulture);
            if(format!=null && Regex.IsMatch(format,"^0{1,32}$")){
                decimal integer;if(!Decimal.TryParse(value,NumberStyles.Float,CultureInfo.InvariantCulture,out integer) || integer!=Decimal.Truncate(integer))throw Bad("Use valores inteiros nas categorias com zeros à esquerda, ou converta-as em texto no Excel.");return integer.ToString(format,CultureInfo.InvariantCulture);
            }
            // Deliberately bounded display support: no arbitrary Excel format engine.
            string code=format==null?"":format.ToLowerInvariant();CultureInfo culture=CultureInfo.CurrentCulture;string dotnet=null;
            if(format==null && style.Id>=14 && style.Id<=22){string[] builtins={culture.DateTimeFormat.ShortDatePattern,"d-MMM-yy","d-MMM","MMM-yy","h:mm tt","h:mm:ss tt","H:mm","H:mm:ss",culture.DateTimeFormat.ShortDatePattern+" H:mm"};dotnet=builtins[style.Id-14];}
            else switch(code){case "mmm/yyyy":dotnet="MMM/yyyy";break;case "dd/mm/yyyy":dotnet="dd/MM/yyyy";break;case "mm-dd-yy":dotnet="MM-dd-yy";break;case "d-mmm-yy":dotnet="d-MMM-yy";break;case "d-mmm":dotnet="d-MMM";break;case "mmm-yy":dotnet="MMM-yy";break;case "h:mm am/pm":dotnet="h:mm tt";break;case "h:mm:ss am/pm":dotnet="h:mm:ss tt";break;case "h:mm":dotnet="H:mm";break;case "h:mm:ss":dotnet="H:mm:ss";break;case "m/d/yy h:mm":dotnet="M/d/yy H:mm";break;}
            if(dotnet==null)throw Bad("Formato da categoria não suportado na atualização XLSX. Use datas padrão, mmm/yyyy, dd/MM/yyyy ou zeros inteiros (000000), ou converta a categoria em texto no Excel e salve.");
            if(number<0 || (!date1904 && Math.Floor(number)==60))throw Bad("Data de categoria XLSX inválida. Converta esta categoria em texto no Excel.");
            if(dotnet.IndexOf("tt",StringComparison.Ordinal)>=0){culture=(CultureInfo)culture.Clone();culture.DateTimeFormat.AMDesignator="AM";culture.DateTimeFormat.PMDesignator="PM";}
            try{DateTime date=DateTime.FromOADate(date1904?number+1462:number<60?number+1:number);string label=date.ToString(dotnet,culture);string month=date.ToString("MMM",culture);if(month.EndsWith(".",StringComparison.Ordinal))label=label.Replace(month,month.TrimEnd('.'));return label;}catch(ArgumentException){throw Bad("Data de categoria fora dos limites. Converta a categoria em texto no Excel.");}
        }
        static string Value(XElement cell,XNamespace ns,List<string> shared)
        {
            XElement cached=cell.Element(ns+"v");string kind=Attr(cell,"t");
            if(kind=="e")throw Bad("A seleção contém erro de célula. Corrija-o no Excel e salve antes de importar.");
            if(cell.Element(ns+"f")!=null && (cached==null || String.IsNullOrWhiteSpace(cached.Value)))throw Bad("A seleção contém fórmula sem valor em cache. Recalcule e salve no Excel antes de importar.");
            string value=cached==null?"":cached.Value;
            if(kind=="inlineStr")value=RichText(cell.Element(ns+"is"),ns);
            else if(kind=="s"){int index;if(!Int32.TryParse(value,NumberStyles.None,CultureInfo.InvariantCulture,out index) || index<0 || index>=shared.Count)throw Bad("Índice de texto compartilhado inválido.");value=shared[index];}
            else if(kind!="" && kind!="n" && kind!="str" && kind!="b" && kind!="d")throw Bad("Tipo de célula XLSX incompatível.");
            if(value.Length>100000)throw Bad("Texto de célula grande demais.");return value;
        }
        public static IList<string> ListSheets(string path){using(Package package=new Package(path))return package.Sheets.Select(s=>Attr(s,"name")).ToList().AsReadOnly();}
        internal static List<List<string>> ReadTable(string path,string sheet,string range,string type,out WorkbookSource link)
        {
            ChartCore.RowLimit(type);Area area=ParseArea(range,type);
            try{using(Package package=new Package(path)){
                XElement selected=package.Sheets.FirstOrDefault(s=>String.Equals(Attr(s,"name"),sheet,StringComparison.OrdinalIgnoreCase));if(selected==null)throw Bad("A planilha selecionada não existe mais no arquivo. Escolha uma aba válida.");
                XAttribute id=selected.Attributes().FirstOrDefault(a=>a.Name.LocalName=="id" && (a.Name.NamespaceName=="http://schemas.openxmlformats.org/officeDocument/2006/relationships" || a.Name.NamespaceName=="http://purl.oclc.org/ooxml/officeDocument/relationships"));
                XElement relation=id==null?null:package.Relations.FirstOrDefault(r=>Attr(r,"Id")==id.Value && Attr(r,"Type").EndsWith("/worksheet",StringComparison.Ordinal));if(relation==null)throw Bad("A aba selecionada não é uma planilha válida.");
                string sheetPath=PartPath(package.WorkbookPath,Attr(relation,"Target"));List<string> shared=new List<string>();XElement sharedRelation=package.Relations.FirstOrDefault(r=>Attr(r,"Type").EndsWith("/sharedStrings",StringComparison.Ordinal));
                XDocument strings=package.ReadXml(sharedRelation==null?PartPath(package.WorkbookPath,"sharedStrings.xml"):PartPath(package.WorkbookPath,Attr(sharedRelation,"Target")),sharedRelation!=null);
                if(strings!=null){XNamespace sn=SheetNamespace(strings);foreach(XElement item in strings.Root.Elements(sn+"si")){if(shared.Count>=16000)throw Bad("A planilha tem textos compartilhados demais.");string value=RichText(item,sn);if(value.Length>100000)throw Bad("Texto de célula grande demais.");shared.Add(value);}}
                XDocument doc=package.ReadXml(sheetPath,true);XNamespace ns=SheetNamespace(doc);XElement data=doc.Root.Element(ns+"sheetData");if(data==null)throw Bad("A aba selecionada está vazia.");
                List<CellStyle> styles=ReadStyles(package);
                Dictionary<int,Dictionary<int,string>> grid=new Dictionary<int,Dictionary<int,string>>();Dictionary<long,int> numericStyles=new Dictionary<long,int>();HashSet<long> numericCells=new HashSet<long>(),percentCells=new HashSet<long>();HashSet<int> rowIds=new HashSet<int>();int previousRow=0,count=0;Area used=new Area{Top=1048577,Left=16385,Bottom=0,Right=0};
                foreach(XElement rowElement in data.Elements(ns+"row")){
                    int row=previousRow+1;string rowRef=Attr(rowElement,"r");if(rowRef!="" && !Int32.TryParse(rowRef,NumberStyles.None,CultureInfo.InvariantCulture,out row))throw Bad("Linha XLSX inválida.");
                    if(row<1 || row>1048576 || !rowIds.Add(row))throw Bad("Linha XLSX fora dos limites ou duplicada.");previousRow=row;HashSet<int> columns=new HashSet<int>();int previousCol=0;
                    foreach(XElement cell in rowElement.Elements(ns+"c")){
                        if(++count>MaxCells)throw Bad("A aba contém mais de 16.384 células. Salve uma cópia menor para importar.");int col=previousCol+1;string reference=Attr(cell,"r");if(reference!=""){int cellRow;col=Coordinate(reference,out cellRow,false);if(cellRow!=row)throw Bad("Referência de célula não corresponde à linha.");}
                        if(col>16384 || !columns.Add(col))throw Bad("Coluna XLSX fora dos limites ou duplicada.");previousCol=col;
                        if(area!=null && !area.Contains(row,col))continue;string value=Value(cell,ns,shared);
                        if(!grid.ContainsKey(row))grid.Add(row,new Dictionary<int,string>());grid[row].Add(col,value);
                        if(Attr(cell,"t")=="" || Attr(cell,"t")=="n"){
                            long key=row*16384L+col;numericCells.Add(key);int style=0;string raw=Attr(cell,"s");if((raw!="" && !Int32.TryParse(raw,NumberStyles.None,CultureInfo.InvariantCulture,out style)) || style>=styles.Count)throw Bad("Índice de estilo XLSX inválido.");numericStyles.Add(key,style);if(styles[style].Percent)percentCells.Add(key);
                        }
                        if(!String.IsNullOrWhiteSpace(value)){used.Top=Math.Min(used.Top,row);used.Bottom=Math.Max(used.Bottom,row);used.Left=Math.Min(used.Left,col);used.Right=Math.Max(used.Right,col);}
                    }
                }
                if(used.Bottom==0)throw Bad("O intervalo selecionado está vazio. Inclua cabeçalho e dados salvos.");if(area==null){area=used;Bound(area,type);}
                List<List<string>> table=new List<List<string>>();for(int row=area.Top;row<=area.Bottom;row++){List<string> cells=new List<string>();Dictionary<int,string> values;grid.TryGetValue(row,out values);for(int col=area.Left;col<=area.Right;col++){string value="";if(values!=null)values.TryGetValue(col,out value);cells.Add(value??"");}table.Add(cells);}
                if(table.Count>1)for(int col=0;col<table[0].Count;col++)if(CategoryHeader(table[0][col]))for(int row=1;row<table.Count;row++){
                    int style;if(numericStyles.TryGetValue((area.Top+row)*16384L+area.Left+col,out style))table[row][col]=CategoryValue(table[row][col],styles[style],package.Date1904);
                }
                // Excel's 1904 date system is workbook-scoped. Shift only numeric date cells,
                // leaving text dates and progress/numeric series untouched.
                if(type=="gantt" && package.Date1904 && table.Count>1)for(int col=0;col<table[0].Count;col++)if(DateHeader(table[0][col]))for(int row=1;row<table.Count;row++){
                    double serial;if(numericCells.Contains((area.Top+row)*16384L+area.Left+col) && Double.TryParse(table[row][col],NumberStyles.Float,CultureInfo.InvariantCulture,out serial))table[row][col]=(serial+1462).ToString("R",CultureInfo.InvariantCulture);
                }
                if(type=="gantt" && table.Count>1)for(int col=0;col<table[0].Count;col++)if(Header(table[0][col])=="progresso" || Header(table[0][col])=="progress")for(int row=1;row<table.Count;row++){
                    double value;if(percentCells.Contains((area.Top+row)*16384L+area.Left+col) && Double.TryParse(table[row][col],NumberStyles.Float,CultureInfo.InvariantCulture,out value))table[row][col]=(value*100).ToString("R",CultureInfo.InvariantCulture);
                }
                link=new WorkbookSource{path=package.Path,sheet=Attr(selected,"name"),range=area.Reference};return table;
            }}catch(XmlException e){throw new InvalidDataException("XML XLSX inválido: "+e.Message,e);}catch(NotSupportedException e){throw new InvalidDataException("Formato XLSX incompatível.",e);}
        }
        public static ChartDocument Import(string path,string sheet,string range,string type)
        {
            WorkbookSource link;List<List<string>> table=ReadTable(path,sheet,range,type,out link);return FromTable(table,type,link);
        }
        internal static ChartDocument FromTable(List<List<string>> table,string type,WorkbookSource link){ChartDocument result=ChartCore.ImportTable(table,type,true);result.excelLink=link;string source="Excel · "+link.sheet+" · "+link.range;result.source=source.Length<=100?source:source.Substring(0,100);return ChartCore.Validate(result);}
        public static ChartDocument Refresh(ChartDocument current)
        {
            ChartDocument result=ChartCore.Clone(current);if(result.excelLink==null)throw new ArgumentException("Este gráfico não possui um vínculo XLSX salvo.");WorkbookSource source=result.excelLink;
            ChartDocument imported=Import(source.path,source.sheet,source.range,result.type);result.rows=imported.rows;result.excelLink=imported.excelLink;return ChartCore.Validate(result);
        }
    }
}
