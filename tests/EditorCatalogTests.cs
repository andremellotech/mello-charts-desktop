using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using MelloCharts;

public static class EditorCatalogTests {
  static int passed,failed;
  static readonly BindingFlags Private=BindingFlags.Instance|BindingFlags.NonPublic;
  static void Is(bool condition,string message){if(!condition)throw new Exception(message);}
  static void Test(string name,Action test){try{test();passed++;Console.WriteLine("PASS "+name);}catch(Exception ex){failed++;Console.WriteLine("FAIL "+name+": "+(ex.InnerException??ex).Message);}}
  static T Field<T>(object owner,string name){return (T)owner.GetType().GetField(name,Private).GetValue(owner);}
  static object Call(object owner,string name,params object[] args){return owner.GetType().GetMethod(name,Private).Invoke(owner,args);}
  static ChartDocument Read(EditorForm editor){return (ChartDocument)Call(editor,"ReadEditor");}
  [STAThread] public static int Main(){
    foreach(ChartDefinition entry in ChartCatalog.All){ChartDefinition definition=entry;
      Test("editor roundtrip "+definition.Key,delegate{
        ChartDocument initial=ChartCore.Example(definition.Key);
        using(EditorForm editor=new EditorForm(null,initial,null)){
          Is(ChartCore.Serialize(initial)==ChartCore.Serialize(Read(editor)),"Editor mudou os dados ou opções iniciais");Is(!editor.HasPendingChanges,"Exemplo inicial marcado como alterado");
          Is(Field<Label>(editor,"typeLabel").Text==definition.Name.ToUpperInvariant(),"Nome incorreto");Is(Field<Label>(editor,"method").Text==definition.Method,"Método incorreto");
          DataGridView grid=Field<DataGridView>(editor,"grid");Is(grid.Columns.Count==(definition.SingleValue?2:definition.Key=="waterfall"||definition.Key=="scatter"?3:4),"Esquema incorreto");
          // Hidden parent forms report Visible=false even for visible children; inspect the local state after showing invisibly.
          editor.ShowInTaskbar=false;editor.Opacity=0;editor.Show();Application.DoEvents();
          Is(Field<CheckBox>(editor,"baseline").Visible==definition.Baseline,"Baseline indevido");Is(Field<CheckBox>(editor,"total").Visible==definition.Total,"Total indevido");
          Is(Field<CheckBox>(editor,"gridLines").Visible==(definition.Key!="pie"&&definition.Key!="donut"),"Opção de grade incorreta");
        }
      });
      Test("data import preserves appearance "+definition.Key,delegate{
        using(EditorForm editor=new EditorForm(null,ChartCore.Example(definition.Key),null)){
          Field<ComboBox>(editor,"style").SelectedIndex=2;Field<ComboBox>(editor,"palette").SelectedIndex=1;Field<ComboBox>(editor,"decimals").SelectedIndex=2;Field<CheckBox>(editor,"compact").Checked=true;Field<CheckBox>(editor,"gridLines").Checked=false;Field<CheckBox>(editor,"labels").Checked=false;Field<CheckBox>(editor,"baseline").Checked=true;Field<CheckBox>(editor,"total").Checked=true;
          ChartDocument next=ChartCore.ImportDelimited(ChartCore.ToCsv(ChartCore.Example(definition.Key)),definition.Key);Call(editor,"ImportData",next,"Teste de importação");ChartDocument read=Read(editor);
          Is(read.style=="dark"&&read.palette=="forest"&&read.decimals==2&&read.compactNumbers&&!read.showGrid&&!read.showLabels&&read.showBaseline&&read.includeTotal,"Importação perdeu aparência");Is(read.source=="Teste de importação","Origem ausente");
        }
      });
      if(definition.SingleValue)Test("single-value precision and row insertion "+definition.Key,delegate{
        ChartDocument initial=ChartCore.Example(definition.Key);initial.rows[0].value=.30000000000000004;
        using(EditorForm editor=new EditorForm(null,initial,null)){
          Is(Read(editor).rows[0].value==initial.rows[0].value,"Precisão perdida");Call(editor,"AddRow");ChartDocument read=Read(editor);Is(read.rows.Count==initial.rows.Count+1&&read.rows.Last().value==0,"Nova categoria inválida");
          DataGridView grid=Field<DataGridView>(editor,"grid");grid.Rows[0].Cells["value"].Value="-1";Is(editor.HasPendingChanges,"Valor inválido descartado");Is(!Field<Button>(editor,"apply").Enabled,"Valor negativo permitiu aplicar");
        }
      });
    }
    Test("gallery has all distinct models and chooses the selected model",delegate{
      using(ChartGalleryForm gallery=new ChartGalleryForm()){
        IList models=Field<IList>(gallery,"models");Is(models.Count==14&&ChartCatalog.Contains(gallery.SelectedType),"Galeria incompleta");
        object last=models[models.Count-1];Call(gallery,"SelectCard",last);Is(gallery.SelectedType=="bubble","Seleção incorreta");Call(gallery,"Choose");Is(gallery.DialogResult==DialogResult.OK,"Seleção não aceita");
      }
    });
    Test("PNG renderer fills ellipse and closed polygon geometry",delegate{
      Scene scene=new Scene{Width=100,Height=100};scene.Elements.Add(new SceneElement{Type="ellipse",Color="#ff0000",X=10,Y=10,W=30,H=30});scene.Elements.Add(new SceneElement{Type="polygon",Color="#0000ff",Points=new double[]{50,50,90,50,70,90}});
      using(Bitmap bitmap=ChartPreview.RenderImage(scene,100,100)){
        Is(bitmap.GetPixel(25,25).R>240&&bitmap.GetPixel(25,25).B<10,"Elipse ausente");Is(bitmap.GetPixel(70,65).B>240&&bitmap.GetPixel(70,65).R<10,"Polígono ausente");Is(bitmap.GetPixel(10,10).R>240&&bitmap.GetPixel(10,10).G>240,"Elipse virou retângulo");
        using(MemoryStream stream=new MemoryStream()){bitmap.Save(stream,System.Drawing.Imaging.ImageFormat.Png);Is(stream.Length>100,"PNG inválido");}
      }
    });
    Console.WriteLine(passed+" editor/catalog checks passed; "+failed+" failed");return failed==0?0:1;
  }
}
