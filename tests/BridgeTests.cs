using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using MelloCharts;
public static class BridgeTests {
  private static int count=0;
  private static void Check(bool value,string name){if(!value)throw new Exception(name);count++;Console.WriteLine("OK "+name);}
  [STAThread]public static int Main(){try{
    ChartDocument d=ChartCore.Example("buildup");d.title="Água e energia · PEP 1";d.rows[0].label="São José <&>";d.style="dark";d.compactNumbers=true;d.showGrid=false;
    IList<string> chunks=Metadata.Encode(d);Check(chunks.All(s=>s.Length<=200&&s.All(c=>"0123456789ABCDEF".Contains(c))),"Tags em hexadecimal preservam texto após conversão em maiúsculas do Office");
    Check(ChartCore.Serialize(Metadata.Decode(chunks.Select(s=>s.ToUpperInvariant()).ToList()))==ChartCore.Serialize(d),"Backup no grupo preserva acentos, caixa e todos os valores");
    bool rejected=false;try{Metadata.Decode(new List<string>{"XYZ"});}catch{rejected=true;}Check(rejected,"Metadados corrompidos são rejeitados");
    XDocument xml=XDocument.Parse(new Connect().GetCustomUI("Microsoft.PowerPoint.Presentation"));XNamespace ns="http://schemas.microsoft.com/office/2009/07/customui";
    var controls=xml.Descendants(ns+"button").ToList();Check(controls.Count==8,"Ribbon possui criação, edição, duplicação, ajuste e suporte");
    foreach(var control in controls){string method=(string)control.Attribute("onAction");Check(typeof(Connect).GetMethod(method)!=null,"Callback disponível: "+method);}
    foreach(Type t in new[]{typeof(IDTExtensibility2),typeof(IRibbonExtensibility)}){var a=(InterfaceTypeAttribute)Attribute.GetCustomAttribute(t,typeof(InterfaceTypeAttribute));Check(a.Value==ComInterfaceType.InterfaceIsDual,"ABI COM dual: "+t.Name);}
    object connect=new Connect();IntPtr p=Marshal.GetIUnknownForObject(connect);try{foreach(Type t in new[]{typeof(IDTExtensibility2),typeof(IRibbonExtensibility)}){Guid id=t.GUID;IntPtr result;int hr=Marshal.QueryInterface(p,ref id,out result);Check(hr==0&&result!=IntPtr.Zero,"QueryInterface: "+t.Name);if(result!=IntPtr.Zero)Marshal.Release(result);}}finally{Marshal.Release(p);}
    ChartDocument precise=ChartCore.Example("buildup");precise.rows[0].actual=0.30000000000000004;precise.style="dark";precise.compactNumbers=true;precise.showGrid=false;
    using(EditorForm editor=new EditorForm(null,precise,null)){
      MethodInfo read=typeof(EditorForm).GetMethod("ReadEditor",BindingFlags.Instance|BindingFlags.NonPublic);ChartDocument roundTrip=(ChartDocument)read.Invoke(editor,null);
      Check(BitConverter.DoubleToInt64Bits(roundTrip.rows[0].actual)==BitConverter.DoubleToInt64Bits(precise.rows[0].actual),"Abrir o editor mantém precisão numérica antes de qualquer edição");
      Check(roundTrip.style=="dark" && roundTrip.compactNumbers && !roundTrip.showGrid,"Editor preserva o estilo e as opções de leitura");
      ChartDocument imported=ChartCore.ImportDelimited("categoria;realizado;projetado;baseline\nNova;1;2;3","buildup");typeof(EditorForm).GetMethod("PresentationTo",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(editor,new object[]{imported});
      Check(imported.style=="dark" && imported.compactNumbers && !imported.showGrid && imported.rows[0].actual==1,"Importação mantém aparência e usa os novos dados");
    }
    FakeSlide slide=new FakeSlide();ChartDocument nativeDoc=ChartCore.Example("columns");nativeDoc.style="dark";nativeDoc.compactNumbers=true;nativeDoc.showGrid=false;ChartReference first=PowerPointBridge.DrawOnSlide(slide,nativeDoc,null);
    Check(slide.Shapes.Count==1&&slide.Shapes[1].Type==6,"Desenho nativo termina em um grupo");FakeShape original=slide.Shapes[1];original.Left=52;original.Top=36;original.Width=560;original.Height=315;original.Rotation=12;
    nativeDoc.title="Revisão";ChartReference revised=PowerPointBridge.DrawOnSlide(slide,nativeDoc,first);
    Check(slide.Shapes.Count==1&&slide.Shapes[1].Left==52&&slide.Shapes[1].Rotation==12,"Atualização preserva geometria e remove apenas o original");
    Check(Metadata.Read(slide.Shapes[1]).title==nativeDoc.title,"Grupo nativo recupera modelo completo");original=slide.Shapes[1];slide.Shapes.FailNextName=true;
    rejected=false;try{PowerPointBridge.DrawOnSlide(slide,nativeDoc,revised);}catch{rejected=true;}
    Check(rejected&&slide.Shapes.Count==1&&Object.ReferenceEquals(original,slide.Shapes[1]),"Falha ao nomear primeira forma limpa o novo objeto e preserva original");
    slide.Shapes.FailTags=true;rejected=false;try{PowerPointBridge.DrawOnSlide(slide,nativeDoc,revised);}catch{rejected=true;}slide.Shapes.FailTags=false;
    Check(rejected&&slide.Shapes.Count==1&&Object.ReferenceEquals(original,slide.Shapes[1]),"Falha ao gravar dados limpa novo grupo sem excluir o original");
    original.FailDelete=true;rejected=false;try{PowerPointBridge.DrawOnSlide(slide,nativeDoc,revised);}catch(Exception ex){rejected=ex.Message.Contains("anterior");}
    Check(rejected&&slide.Shapes.Count==2,"Falha na exclusão final mantém os dois grupos e informa duplicação");
    Console.WriteLine(count+" verificações da ponte passaram.");return 0;
  }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}}
}
// A small native-shape contract adapter. These tests do not activate PowerPoint.
public sealed class FakeSlide {public int SlideID=42,SlideIndex=1;public FakeShapes Shapes=new FakeShapes();public object Parent=new FakePageOwner();}
public sealed class FakePageOwner {public FakePageSetup PageSetup=new FakePageSetup();}
public sealed class FakePageSetup {public double SlideWidth=720,SlideHeight=540;}
public sealed class FakeShapes {
  public readonly List<FakeShape> Items=new List<FakeShape>();public bool FailNextName,FailTags;private int next;
  public int Count{get{return Items.Count;}}public FakeShape this[int index]{get{return Items[index-1];}}
  internal FakeShape Add(int type,float x,float y,float w,float h){FakeShape s=new FakeShape(this,++next,type,x,y,w,h);Items.Add(s);return s;}
  public FakeShape AddShape(int type,float x,float y,float w,float h){return Add(1,x,y,w,h);}
  public FakeShape AddLine(float x,float y,float x2,float y2){return Add(9,Math.Min(x,x2),Math.Min(y,y2),Math.Abs(x2-x),Math.Abs(y2-y));}
  public FakeShape AddTextbox(int direction,float x,float y,float w,float h){return Add(17,x,y,w,h);}
  public FakeRange Range(string[] names){return new FakeRange(this,names);}
  public sealed class FakeRange {private FakeShapes owner;private string[] names;public FakeRange(FakeShapes o,string[] n){owner=o;names=n;}public FakeShape Group(){var parts=owner.Items.Where(s=>names.Contains(s.Name)).ToArray();if(parts.Length!=names.Length)throw new Exception("Objeto ausente");FakeShape canvas=parts[0];FakeShape group=owner.Add(6,canvas.Left,canvas.Top,canvas.Width,canvas.Height);foreach(FakeShape p in parts)owner.Items.Remove(p);return group;}}
}
public sealed class FakeShape {
  private FakeShapes owner;private string name="";public int Id,Type,LockAspectRatio;public float Left,Top,Rotation;private float width,height;public bool FailNextHeight;public float Width{get{return width;}set{width=value;}}public float Height{get{return height;}set{if(FailNextHeight){FailNextHeight=false;throw new Exception("Redimensionamento bloqueado");}height=value;}}public string AlternativeText,ManualStyle;public bool FailDelete;public object ParentGroup;
  public FakeFill Fill=new FakeFill();public FakeLine Line=new FakeLine();public FakeFrame TextFrame=new FakeFrame();public FakeTags Tags;
  public FakeShape(FakeShapes o,int id,int type,float x,float y,float w,float h){owner=o;Id=id;Type=type;Left=x;Top=y;Width=w;Height=h;Tags=new FakeTags(o);}
  public string Name {get{return name;}set{if(owner.FailNextName){owner.FailNextName=false;throw new Exception("Nome bloqueado");}name=value;}}
  public void Delete(){if(FailDelete)throw new Exception("Exclusão bloqueada");owner.Items.Remove(this);}public void Select(){}
  public FakeDuplicateRange Duplicate(){FakeShape copy=owner.Add(Type,Left,Top,Width,Height);copy.Rotation=Rotation;copy.LockAspectRatio=LockAspectRatio;copy.ManualStyle=ManualStyle;copy.AlternativeText=AlternativeText;copy.Tags=Tags.Clone(owner);return new FakeDuplicateRange(copy);}
}
public sealed class FakeDuplicateRange {private readonly FakeShape shape;public FakeDuplicateRange(FakeShape item){shape=item;}public int Count{get{return 1;}}public FakeShape this[int index]{get{if(index!=1)throw new IndexOutOfRangeException();return shape;}}}
public sealed class FakeFill {public FakeColor ForeColor=new FakeColor();public int Visible;public void Solid(){}}
public sealed class FakeLine {public FakeColor ForeColor=new FakeColor();public int Visible,DashStyle;public float Weight;}
public sealed class FakeColor {public int RGB;}
public sealed class FakeFrame {public float MarginLeft,MarginTop,MarginRight,MarginBottom;public int AutoSize,WordWrap;public FakeRangeText TextRange=new FakeRangeText();}
public sealed class FakeRangeText {public string Text;public FakeFont Font=new FakeFont();public FakeParagraph ParagraphFormat=new FakeParagraph();}
public sealed class FakeFont {public string Name;public float Size;public int Bold;public FakeColor Color=new FakeColor();}
public sealed class FakeParagraph {public int Alignment;}
public sealed class FakeTags {
  private FakeShapes owner;private Dictionary<string,string> values=new Dictionary<string,string>();public FakeTags(FakeShapes o){owner=o;}
  public void Add(string name,string value){if(owner.FailTags)throw new Exception("Tags bloqueadas");values[name.ToUpperInvariant()]=value.ToUpperInvariant();}
  public string this[string key]{get{string result;return values.TryGetValue(key.ToUpperInvariant(),out result)?result:"";}}
  public FakeTags Clone(FakeShapes target){FakeTags copy=new FakeTags(target);foreach(var entry in values)copy.values.Add(entry.Key,entry.Value);return copy;}
}
