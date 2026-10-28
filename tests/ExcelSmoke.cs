using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using MelloCharts;

public static class ExcelSmoke {
  static void Require(bool condition,string message){if(!condition)throw new Exception(message);}
  [STAThread]public static int Main(string[] args){
    if(args.Length!=1){Console.Error.WriteLine("Informe uma pasta existente para o XLSX de teste. Não instala o suplemento.");return 2;}
    string folder=Path.GetFullPath(args[0]);if(!Directory.Exists(folder))return 2;
    Process[] processes=Process.GetProcessesByName("EXCEL");bool alreadyRunning=processes.Length>0;foreach(Process process in processes)process.Dispose();
    object application=null,workbooks=null,workbook=null,worksheets=null,sheet=null,range=null,ganttSheet=null,ganttRange=null,categorySheet=null,categoryRange=null;int result=1;
    try{
      application=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true));dynamic app=application;
      workbooks=(object)app.Workbooks;dynamic books=workbooks;workbook=(object)books.Add();dynamic book=workbook;worksheets=(object)book.Worksheets;dynamic sheets=worksheets;
      if((int)sheets.Count<2){object first=null;try{first=(object)sheets[1];sheet=(object)sheets.Add(Type.Missing,first,1,Type.Missing);}finally{Release(first);}}
      else sheet=(object)sheets[2];dynamic dataSheet=sheet;dataSheet.Name="Dados salvos";dataSheet.Activate();
      range=(object)dataSheet.Range["H500:K502"];dynamic table=range;
      table.Value2=new object[,]{{"categoria","realizado","projetado","baseline"},{"Equipamentos",0.0,2.75,8.0},{"Serviços",9.5,1.25,12.0}};
      Cell(sheet,"I501","=1+2.125",null,true);table.Calculate();table.Select();
      string output=Path.Combine(folder,"mello-charts-excel-smoke-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N").Substring(0,8)+".xlsx");
      book.SaveAs(output,51);
      ChartDocument captured=ExcelBridge.CaptureSelection(application,"columns");
      Require(captured.rows.Count==2&&captured.rows[0].actual==3.125&&captured.rows[0].forecast==2.75,"Captura de dados ou fórmula incorreta");
      Require(captured.excelLink!=null&&captured.excelLink.sheet=="Dados salvos"&&captured.excelLink.range=="$H$500:$K$502"&&String.Equals(captured.excelLink.path,output,StringComparison.OrdinalIgnoreCase),"Vínculo da segunda aba/intervalo deslocado incorreto");
      Console.WriteLine("OK Excel: seleção H500:K502 da segunda aba, fórmula calculada e decimais preservados.");
      captured.title="Título mantido ao atualizar";captured.seriesActual="Confirmado";captured.decimals=2;
      Cell(sheet,"I501","=10+2.375",null,true);table.Calculate();book.Save();
      ChartDocument refreshed=WorkbookData.Refresh(captured);
      Require(refreshed.rows[0].actual==12.375&&captured.rows[0].actual==3.125&&refreshed.title=="Título mantido ao atualizar"&&refreshed.seriesActual=="Confirmado"&&refreshed.decimals==2,"Atualização dos dados salvos perdeu dados/estilo ou alterou o original");
      Console.WriteLine("OK Excel: XLSX salvo atualizado manualmente, fórmula em cache12.375 e aparência preservada.");
      categorySheet=(object)sheets.Add(Type.Missing,sheet,1,Type.Missing);dynamic categories=categorySheet;categories.Name="Categorias formatadas";categories.Activate();
      categoryRange=(object)categories.Range["H500:I502"];dynamic categoryTable=categoryRange;
      categoryTable.Value2=new object[,]{{"categoria","valor"},{45678.0,10.0},{123.0,20.0}};categoryTable.ColumnWidth=20;
      MonthYearFormat(categorySheet,"H501");Cell(categorySheet,"H502",null,"000000",false);categoryTable.Select();book.Save();
      ChartDocument formattedLive=ExcelBridge.CaptureSelection(application,"pie");
      Console.WriteLine("LIVE Excel categorias: "+formattedLive.rows[0].label+" / "+formattedLive.rows[1].label);
      Require(String.Equals(formattedLive.rows[0].label,"jan/2025",StringComparison.OrdinalIgnoreCase)&&formattedLive.rows[1].label=="000123","Excel não preservou a categoria com data ou código formatado");
      ChartDocument formattedSaved=WorkbookData.Refresh(formattedLive);
      Require(formattedSaved.rows[0].label==formattedLive.rows[0].label&&formattedSaved.rows[1].label==formattedLive.rows[1].label,"Atualização do XLSX mudou categorias formatadas: "+formattedLive.rows[0].label+" / "+formattedSaved.rows[0].label+"; "+formattedLive.rows[1].label+" / "+formattedSaved.rows[1].label);
      Console.WriteLine("OK Excel: categorias formatadas "+formattedLive.rows[0].label+" / "+formattedLive.rows[1].label+" preservadas no XLSX salvo.");
      Cell(categorySheet,"I501","=1/0",null,true);categoryTable.Calculate();RejectCapture(application,"pie","Fórmula Excel =1/0 aceita como valor numérico");
      Cell(categorySheet,"I501",1000000001.0,null,false);categoryTable.Calculate();RejectCapture(application,"pie","Número acima de1bilhão aceito na seleção Excel");
      Cell(categorySheet,"I501",10.0,null,false);categoryTable.Calculate();book.Save();
      Console.WriteLine("OK Excel: erro nativo de fórmula e limite numérico rejeitados.");
      book.Date1904=true;ganttSheet=(object)sheets.Add(Type.Missing,sheet,1,Type.Missing);dynamic timeline=ganttSheet;timeline.Name="Cronograma1904";timeline.Activate();
      ganttRange=(object)timeline.Range["H500:K502"];dynamic gantt=ganttRange;
      gantt.Value2=new object[,]{{"atividade","inicio","fim","progresso"},{"Etapa A",0.0,1.0,0.5},{"Etapa B",2.0,3.0,0.5}};
      Cell(ganttSheet,"I501:J502",null,"yyyy-mm-dd",false);Cell(ganttSheet,"K501",null,"0%",false);Cell(ganttSheet,"K502",null,"0.00",false);gantt.Select();book.Save();
      ChartDocument ganttLive=ExcelBridge.CaptureSelection(application,"gantt"),ganttSaved=WorkbookData.Refresh(ganttLive);
      foreach(ChartDocument doc in new[]{ganttLive,ganttSaved})Require(doc.rows[0].startDate=="1904-01-01"&&doc.rows[0].endDate=="1904-01-02"&&doc.rows[0].progress==50&&doc.rows[1].progress==0.5,"Datas1904 ou porcentagem divergentes entre Excel e XLSX salvo");
      Console.WriteLine("OK Excel: Gantt1904 e progresso50% /0.5 sem formatação coincidem na captura e no XLSX salvo.");
      Console.WriteLine(output);result=0;
    }catch(Exception ex){Console.Error.WriteLine(ex);}
    finally{
      Release(ganttRange);Release(ganttSheet);Release(categoryRange);Release(categorySheet);Release(range);Release(sheet);Release(worksheets);
      if(workbook!=null)try{dynamic book=workbook;book.Close(false);}catch(Exception ex){Console.Error.WriteLine("Falha ao fechar o workbook de teste: "+ex.Message);result=1;}
      Release(workbook);Release(workbooks);
      if(application!=null&&!alreadyRunning)try{dynamic app=application;app.Quit();}catch(Exception ex){Console.Error.WriteLine("Falha ao encerrar o Excel do teste: "+ex.Message);result=1;}
      Release(application);
    }
    return result;
  }
  static void RejectCapture(object application,string type,string message){bool rejected=false;try{ExcelBridge.CaptureSelection(application,type);}catch(ArgumentException){rejected=true;}catch(InvalidOperationException){rejected=true;}Require(rejected,message);}
  static void MonthYearFormat(object worksheet,string address){
    object cell=null;try{dynamic sheet=worksheet;cell=(object)sheet.Range[address];dynamic target=cell;
      // NumberFormatLocal uses the installed Excel language. Verify that the
      // year is rendered, rather than persisting escaped literal y characters.
      foreach(string format in new[]{"mmm/yyyy","mmm/aaaa"}){
        target.NumberFormatLocal=format;string text=Convert.ToString((object)target.Text);
        Console.WriteLine("FORMAT Excel: local="+Convert.ToString((object)target.NumberFormatLocal)+", invariant="+Convert.ToString((object)target.NumberFormat)+", Text="+text);
        if(text.IndexOf("2025",StringComparison.Ordinal)>=0)return;
      }
      throw new Exception("Não foi possível definir um formato de mês/ano que exiba2025 no idioma instalado do Excel.");
    }finally{Release(cell);}
  }
  static void Cell(object worksheet,string address,object value,string format,bool formula){object cell=null;try{dynamic sheet=worksheet;cell=(object)sheet.Range[address];dynamic target=cell;if(value!=null){if(formula)target.Formula=value;else target.Value2=value;}if(format!=null)target.NumberFormat=format;}finally{Release(cell);}}
  // Cleanup stays statically typed after Close/Quit, including IsComObject.
  static void Release(object value){if(value==null)return;try{if(Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}catch(InvalidComObjectException){}catch(COMException){}}
}
