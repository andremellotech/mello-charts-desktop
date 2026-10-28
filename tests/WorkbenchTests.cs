using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using MelloCharts;

public static class WorkbenchTests {
  static int passed,failed;
  static void Is(bool ok,string message){if(!ok)throw new Exception(message);}
  static void Test(string name,Action run){try{run();passed++;Console.WriteLine("PASS "+name);}catch(Exception e){failed++;Console.WriteLine("FAIL "+name+": "+e.GetBaseException().Message);}}
  static object Call(object obj,string name,params object[] args){return obj.GetType().GetMethod(name,BindingFlags.NonPublic|BindingFlags.Instance).Invoke(obj,args);}
  static T Field<T>(object obj,string name){return (T)obj.GetType().GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(obj);}
  [STAThread]public static int Main(){
    Test("optional series names and typography survive JSON and scene",delegate{
      string json=ChartCore.Serialize(ChartCore.Example("lines"));
      json=System.Text.RegularExpressions.Regex.Replace(json,"\"seriesActual\":\"[^\"]*\"", "\"seriesActual\":\"Receita\"");json=System.Text.RegularExpressions.Regex.Replace(json,"\"seriesForecast\":\"[^\"]*\"", "\"seriesForecast\":\"Meta\"");json=System.Text.RegularExpressions.Regex.Replace(json,"\"fontFamily\":\"[^\"]*\"", "\"fontFamily\":\"Calibri\"");
      ChartDocument d=ChartCore.Deserialize(json);string saved=ChartCore.Serialize(d);
      Is(saved.Contains("Receita")&&saved.Contains("Calibri"),"Personalização descartada no projeto");
      var texts=ChartCore.BuildScene(d).Elements.Where(e=>e.Type=="text").Select(e=>e.Text).ToList();
      Is(texts.Any(t=>t=="Receita")&&texts.Any(t=>t.StartsWith("Meta\n")),"Legenda não usa nomes das séries");
    });
    Test("undo and redo restore values without changing destination",delegate{
      using(EditorForm f=new EditorForm(null,ChartCore.Example("columns"),null)){
        DataGridView grid=Field<DataGridView>(f,"grid");double first=((ChartDocument)Call(f,"ReadEditor")).rows[0].actual;
        grid.Rows[0].Cells["actual"].Value="999";
        Is(((ChartDocument)Call(f,"ReadEditor")).rows[0].actual==999,"Edição não aplicada");
        Call(f,"UndoEdit");Is(((ChartDocument)Call(f,"ReadEditor")).rows[0].actual==first,"Desfazer não restaurou valor");
        Call(f,"RedoEdit");Is(((ChartDocument)Call(f,"ReadEditor")).rows[0].actual==999,"Refazer não restaurou valor");
      }
    });
    Test("invalid new edit can be undone while preserving last valid preview",delegate{
      using(EditorForm f=new EditorForm(null,ChartCore.Example("columns"),null)){
        Field<DataGridView>(f,"grid").Rows[0].Cells["actual"].Value="não é número";
        Is(!Field<Button>(f,"apply").Enabled,"Dados inválidos permitiram aplicar");
        Is(Field<ChartPreview>(f,"preview").Scene!=null,"Prévia válida foi apagada");
        Call(f,"UndoEdit");Is(((ChartDocument)Call(f,"ReadEditor")).rows.Count==4,"Não recuperou edição inválida");
      }
    });
    Test("new model editors preserve meaningful columns",delegate{
      foreach(string type in new[]{"gantt","scatter","bubble"})using(EditorForm f=new EditorForm(null,ChartCore.Example(type),null)){
        ChartDocument read=(ChartDocument)Call(f,"ReadEditor");Is(ChartCore.Serialize(read)==ChartCore.Serialize(ChartCore.Example(type)),"Modelo alterado ao abrir");
        Call(f,"AddRow");ChartCore.Validate((ChartDocument)Call(f,"ReadEditor"));
      }
    });
    Test("Gantt appended tasks follow current dates and handle final calendar day",delegate{
      using(EditorForm f=new EditorForm(null,ChartCore.Example("gantt"),null)){
        DataGridView grid=Field<DataGridView>(f,"grid");grid.Rows[grid.Rows.Count-1].Cells["endDate"].Value="2030-01-01";
        Call(f,"AddRow");Call(f,"AddRow");ChartDocument d=(ChartDocument)Call(f,"ReadEditor");
        Is(d.rows[d.rows.Count-2].startDate=="2030-01-01"&&d.rows.Last().startDate=="2030-01-08","Nova tarefa ignorou a última data editada");
        grid.Rows[grid.Rows.Count-1].Cells["endDate"].Value="9999-12-31";Call(f,"AddRow");d=(ChartDocument)Call(f,"ReadEditor");Is(d.rows.Last().endDate=="9999-12-31","Fim do calendário não tratado");
      }
    });
    Test("eighty labeled points fit native chart metadata",delegate{
      ChartDocument d=ChartCore.Example("bubble");d.rows=Enumerable.Range(0,80).Select(i=>new ChartRow{label=new string('漢',32),x=-1000000000,y=1000000000,size=1000000000}).ToList();
      Is(Metadata.Decode(Metadata.Encode(d)).rows.Count==80,"Metadados perderam pontos aceitos no editor");
    });
    Console.WriteLine(passed+" workbench checks passed; "+failed+" failed");return failed==0?0:1;
  }
}
