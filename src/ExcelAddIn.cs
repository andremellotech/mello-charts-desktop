using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MelloCharts {
  // Reads only the user's explicit selection. Formula results come from Value2;
  // category labels use displayed text so dates do not become Excel serials.
  public static class ExcelBridge {
    public static ChartDocument CaptureSelection(object application,string type){
      ChartCatalog.Get(type);object selection=null,cells=null,worksheet=null,workbook=null;
      try{
        dynamic app=application;if(app==null)throw new InvalidOperationException("O Excel não conectou o suplemento.");
        selection=(object)app.Selection;dynamic range=selection;
        if(range==null||Count((object)range.Areas)!=1)throw new InvalidOperationException("Selecione uma tabela contínua, incluindo os cabeçalhos.");
        int rows=Count((object)range.Rows),columns=Count((object)range.Columns);
        if(rows<2||rows>ChartCore.RowLimit(type)+1)throw new InvalidOperationException("Selecione o cabeçalho e de 1 a "+ChartCore.RowLimit(type)+" categorias para este modelo.");
        if(columns<2||columns>64)throw new InvalidOperationException("Selecione somente as colunas da tabela, incluindo os cabeçalhos.");
        worksheet=(object)range.Worksheet;dynamic sheet=worksheet;workbook=(object)sheet.Parent;dynamic book=workbook;
        bool date1904=type=="gantt"&&Convert.ToBoolean((object)book.Date1904,CultureInfo.InvariantCulture);
        object raw=(object)range.Value2;Array values=raw as Array;
        if(values==null||values.Rank!=2||values.GetLength(0)!=rows||values.GetLength(1)!=columns)throw new InvalidOperationException("A seleção não é uma tabela de células válida.");
        cells=(object)range.Cells;dynamic cellRange=cells;List<List<string>> table=new List<List<string>>();int firstRow=values.GetLowerBound(0),firstColumn=values.GetLowerBound(1);
        for(int r=0;r<rows;r++){
          List<string> row=new List<string>();
          for(int c=0;c<columns;c++){
            object value=values.GetValue(r+firstRow,c+firstColumn);
            if(value is ErrorWrapper)throw new InvalidOperationException("A seleção contém um erro do Excel na linha "+(r+1)+". Corrija a fórmula antes de criar o gráfico.");
            string text=value==null?"":Convert.ToString(value,CultureInfo.InvariantCulture);
            string header=r>0?Normalize(table[0][c]):"";
            if(r>0&&type=="gantt"&&date1904&&IsDate(header)&&IsNumeric(value))text=(Convert.ToDouble(value,CultureInfo.InvariantCulture)+1462).ToString("R",CultureInfo.InvariantCulture);
            // Only display category labels. Numeric values retain full precision.
            bool category=r>0&&IsCategory(header),progress=r>0&&type=="gantt"&&(header=="progresso"||header=="progress")&&IsNumeric(value);
            if(category||progress){object cell=null;try{cell=(object)cellRange[r+1,c+1];dynamic displayed=cell;
              if(category){string label=Convert.ToString((object)displayed.Text,CultureInfo.CurrentCulture);if(!String.IsNullOrWhiteSpace(label)&&label.Trim('#').Length>0)text=label;}
              else if(PercentFormat(Convert.ToString((object)displayed.NumberFormat,CultureInfo.InvariantCulture)))text=(Convert.ToDouble(value,CultureInfo.InvariantCulture)*100).ToString("R",CultureInfo.InvariantCulture);
            }finally{Release(cell);}}
            row.Add(text);
          }
          table.Add(row);
        }
        ChartDocument doc=ChartCore.ImportTable(table,type,true);doc.source="Excel · "+Convert.ToString((object)sheet.Name,CultureInfo.InvariantCulture);
        string folder=Convert.ToString((object)book.Path,CultureInfo.InvariantCulture),path=Convert.ToString((object)book.FullName,CultureInfo.InvariantCulture);
        if(!String.IsNullOrWhiteSpace(folder)&&Path.IsPathRooted(path)&&String.Equals(Path.GetExtension(path),".xlsx",StringComparison.OrdinalIgnoreCase)){
          // Range.Address defaults to absolute A1 references and avoids invoking
          // an interop-generated get_Address method through IDispatch.
          string address=Convert.ToString((object)range.Address,CultureInfo.InvariantCulture);
          doc.excelLink=new WorkbookSource{path=path,sheet=Convert.ToString((object)sheet.Name,CultureInfo.InvariantCulture),range=address};
        }
        return ChartCore.Validate(doc);
      }catch(ArgumentException){throw;}catch(InvalidOperationException){throw;}catch(Exception ex){throw new InvalidOperationException("Selecione células de uma única tabela com cabeçalhos. "+ex.Message,ex);}
      finally{Release(workbook);Release(worksheet);Release(cells);Release(selection);}
    }
    private static string Normalize(string text){StringBuilder result=new StringBuilder();foreach(char c in (text??"").Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD))if(CharUnicodeInfo.GetUnicodeCategory(c)!=UnicodeCategory.NonSpacingMark)result.Append(c);return result.ToString();}
    private static bool IsCategory(string header){return header=="categoria"||header=="category"||header=="label"||header=="atividade"||header=="projeto"||header=="pep"||header=="etapa";}
    private static bool IsDate(string header){return header=="inicio"||header=="data inicial"||header=="startdate"||header=="fim"||header=="data final"||header=="enddate";}
    private static bool IsNumeric(object value){return value is double||value is decimal||value is int||value is long||value is float||value is short;}
    private static bool PercentFormat(string format){bool quoted=false,bracket=false;for(int i=0;i<format.Length;i++){char c=format[i];if(c=='\\'||c=='_'||c=='*'){i++;continue;}if(c=='"'){quoted=!quoted;continue;}if(quoted)continue;if(c=='['){bracket=true;continue;}if(c==']'){bracket=false;continue;}if(c=='%'&&!bracket)return true;}return false;}
    private static int Count(object collection){try{dynamic value=collection;return (int)value.Count;}finally{Release(collection);}}
    private static void Release(object value){try{if(value!=null&&Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}catch(InvalidComObjectException){}catch(COMException){}}
  }

  [ComVisible(true),Guid("24C23E45-A813-4CD8-B29A-08F0784C3CF5"),ProgId("MelloCharts.ExcelConnect"),ClassInterface(ClassInterfaceType.AutoDispatch)]
  public sealed class ExcelConnect:IDTExtensibility2,IRibbonExtensibility {
    private object application;private readonly List<EditorForm> editors=new List<EditorForm>();
    public void OnConnection(object app,int mode,object instance,ref Array custom){application=app;}
    public void OnDisconnection(int mode,ref Array custom){CloseEditors();application=null;}
    public void OnAddInsUpdate(ref Array custom){}
    public void OnStartupComplete(ref Array custom){}
    public void OnBeginShutdown(ref Array custom){CloseEditors();}
    public string GetCustomUI(string ribbonId){return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='MelloCharts.Excel.Tab' label='Mello Charts'><group id='MelloCharts.Excel.Create' label='Criar gráfico'><button id='MelloCharts.Excel.Selection' label='Criar da seleção' size='large' imageMso='ChartInsert' onAction='NewFromSelection' screentip='Usar a tabela selecionada' supertip='Inclua os cabeçalhos e selecione uma tabela contínua. Escolha o modelo e revise os dados no editor antes de enviar ao PowerPoint.'/><button id='MelloCharts.Excel.Gallery' label='Galeria de gráficos' size='large' imageMso='ChartType' onAction='NewGallery'/></group><group id='MelloCharts.Excel.Support' label='Suporte'><button id='MelloCharts.Excel.Diagnostics' label='Diagnóstico' imageMso='Info' onAction='ShowDiagnostics'/><button id='MelloCharts.Excel.Help' label='Como usar' imageMso='Help' onAction='ShowHelp'/></group></tab></tabs></ribbon></customUI>";}
    public void NewFromSelection(object control){Open(true);}
    public void NewGallery(object control){Open(false);}
    private void Open(bool fromSelection){try{using(ChartGalleryForm gallery=new ChartGalleryForm()){if(gallery.ShowDialog()!=DialogResult.OK)return;ChartDocument doc=fromSelection?ExcelBridge.CaptureSelection(application,gallery.SelectedType):ChartCore.Example(gallery.SelectedType);EditorForm editor=new EditorForm(null,doc,null);editors.Add(editor);editor.FormClosed+=(s,e)=>{editors.Remove(editor);editor.Dispose();};try{editor.Show();}catch{editors.Remove(editor);editor.Dispose();throw;}}}catch(Exception ex){MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
    private void CloseEditors(){foreach(EditorForm editor in editors.ToArray())if(!editor.IsDisposed){editor.CloseFromHost();editor.Dispose();}editors.Clear();}
    public void ShowDiagnostics(object control){using(SupportForm form=new SupportForm(InstallationHealth.Collect(typeof(ExcelConnect).Assembly.GetName().Version.ToString(3))))form.ShowDialog();}
    public void ShowHelp(object control){MessageBox.Show("1. Selecione uma tabela contínua no Excel, incluindo os cabeçalhos.\n2. Clique em Criar da seleção e escolha o modelo. A galeria mostra os cabeçalhos esperados.\n3. Revise dados, estilo e cores no editor.\n4. Salve o projeto ou use Enviar ao PowerPoint.\n\nAs fórmulas são lidas pelos resultados atuais. O suplemento não altera a planilha. A origem de um XLSX salvo fica registrada no projeto; atualize-a manualmente no editor quando necessário.","Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Information);}
  }
}
