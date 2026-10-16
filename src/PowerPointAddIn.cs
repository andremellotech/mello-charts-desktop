using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

[assembly:AssemblyTitle("Mello Charts Desktop")]
[assembly:AssemblyDescription("Gráficos financeiros editáveis para PowerPoint")]
[assembly:AssemblyCompany("Mello Charts")]
[assembly:AssemblyProduct("Mello Charts")]
[assembly:AssemblyCopyright("Copyright 2026 André Mello")]
[assembly:AssemblyVersion("0.6.0.0")]
[assembly:AssemblyFileVersion("0.6.0.0")]
[assembly:ComVisible(false)]

namespace MelloCharts {
  [ComVisible(true),ComImport,Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),InterfaceType(ComInterfaceType.InterfaceIsDual),TypeLibType(4160)]
  public interface IDTExtensibility2 {
    [DispId(1)]void OnConnection([In,MarshalAs(UnmanagedType.IDispatch)]object application,[In]int connectMode,[In,MarshalAs(UnmanagedType.IDispatch)]object addInInst,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)]ref Array custom);
    [DispId(2)]void OnDisconnection([In]int removeMode,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)]ref Array custom);
    [DispId(3)]void OnAddInsUpdate([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)]ref Array custom);
    [DispId(4)]void OnStartupComplete([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)]ref Array custom);
    [DispId(5)]void OnBeginShutdown([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)]ref Array custom);
  }
  [ComVisible(true),ComImport,Guid("000C0396-0000-0000-C000-000000000046"),InterfaceType(ComInterfaceType.InterfaceIsDual),TypeLibType(4160)]
  public interface IRibbonExtensibility {[DispId(1)][return:MarshalAs(UnmanagedType.BStr)]string GetCustomUI([MarshalAs(UnmanagedType.BStr)]string ribbonId);}

  [ComVisible(true),Guid("49607523-6CCF-4559-97E4-FCF5BE438C3A"),ProgId("MelloCharts.Connect"),ClassInterface(ClassInterfaceType.AutoDispatch)]
  public sealed class Connect:IDTExtensibility2,IRibbonExtensibility {
    private object application,ribbon;private EditorForm editor;private readonly Timer ribbonTimer=new Timer{Interval=1200};private string ribbonState="";private readonly Dictionary<string,object> icons=new Dictionary<string,object>();
    public Connect(){ribbonTimer.Tick+=(s,e)=>RefreshRibbon();}
    public void OnConnection(object app,int mode,object instance,ref Array custom){application=app;if(ribbon!=null)ribbonTimer.Start();RefreshRibbon();}
    public void OnDisconnection(int mode,ref Array custom){ribbonTimer.Stop();if(editor!=null && !editor.IsDisposed)editor.CloseFromHost();editor=null;application=null;ribbon=null;icons.Clear();}
    public void OnAddInsUpdate(ref Array custom){}
    public void OnStartupComplete(ref Array custom){}
    public void OnBeginShutdown(ref Array custom){ribbonTimer.Stop();if(editor!=null && !editor.IsDisposed)editor.CloseFromHost();editor=null;}
    public void OnRibbonLoad(object ui){ribbon=ui;ribbonState="";if(application!=null)ribbonTimer.Start();RefreshRibbon();}
    private void RefreshRibbon(){if(ribbon==null)return;string next=CanCreate(null)+"/"+CanEdit(null);if(next==ribbonState)return;try{dynamic ui=ribbon;ui.Invalidate();ribbonState=next;}catch{}}
    public bool CanCreate(object control){try{PowerPointBridge.CurrentSlide(application);return true;}catch{return false;}}
    public bool CanEdit(object control){try{PowerPointBridge.Selected(application);return true;}catch{return false;}}
    public string GetCustomUI(string ribbonId){return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'><ribbon><tabs><tab id='MelloCharts.Tab' label='Mello Charts'><group id='MelloCharts.Create' label='Novo gráfico'><button id='MelloCharts.Gallery' label='Galeria de gráficos' size='large' getImage='GetImage' getEnabled='CanCreate' onAction='NewGallery' screentip='Explore todos os gráficos' supertip='Escolha um gráfico por objetivo e abra um exemplo editável.'/><button id='MelloCharts.Build' label='Build-up' size='large' getImage='GetImage' getEnabled='CanCreate' onAction='NewBuildUp' screentip='Realizado + projetado restante' supertip='Duas parcelas na mesma coluna, baseline de referência e total opcional.'/><button id='MelloCharts.Waterfall' label='Cascata' size='large' getImage='GetImage' getEnabled='CanCreate' onAction='NewWaterfall' screentip='Origem das variações' supertip='Base, variações positivas e negativas, subtotais e total.'/><button id='MelloCharts.Columns' label='Comparativo' size='large' getImage='GetImage' getEnabled='CanCreate' onAction='NewColumns' screentip='Comparar séries' supertip='Realizado e projetado em colunas, com baseline opcional.'/></group><group id='MelloCharts.Edit' label='Gráfico selecionado'><button id='MelloCharts.EditSelected' label='Editar dados' size='large' getImage='GetImage' getEnabled='CanEdit' onAction='EditSelected' supertip='Selecione o gráfico ou um elemento dele para recuperar seus dados.'/><button id='MelloCharts.Duplicate' label='Duplicar gráfico' getImage='GetImage' getEnabled='CanEdit' onAction='DuplicateChart' supertip='Cria outra cópia editável, mantendo o original.'/><button id='MelloCharts.Fit' label='Ajustar ao slide' getImage='GetImage' getEnabled='CanEdit' onAction='FitChart' supertip='Centraliza e ajusta ao tamanho do slide em proporção 16:9. Remove a rotação.'/></group><group id='MelloCharts.Support' label='Suporte'><button id='MelloCharts.Diagnostics' label='Diagnóstico' getImage='GetImage' onAction='ShowDiagnostics'/><button id='MelloCharts.Help' label='Como usar' getImage='GetImage' onAction='ShowHelp'/></group></tab></tabs></ribbon></customUI>";}
    public object GetImage(object control){
      string id="Help";try{dynamic c=control;id=(string)c.Id;}catch{}object icon;if(icons.TryGetValue(id,out icon))return icon;using(Bitmap image=RibbonIcons.Create(id))icon=PictureHost.Convert(image);icons[id]=icon;return icon;
    }
    public void NewBuildUp(object control){Open("buildup",false);}
    public void NewWaterfall(object control){Open("waterfall",false);}
    public void NewColumns(object control){Open("columns",false);}
    public void NewGallery(object control){
      try{
        if(editor!=null && !editor.IsDisposed){if(editor.WindowState==FormWindowState.Minimized)editor.WindowState=FormWindowState.Normal;editor.BringToFront();editor.Activate();return;}
        PowerPointBridge.CurrentSlide(application);
        using(ChartGalleryForm gallery=new ChartGalleryForm())if(gallery.ShowDialog(new PowerPointOwner(application))==DialogResult.OK)Open(gallery.SelectedType,false);
      }catch(Exception ex){MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
    }
    public void EditSelected(object control){Open(null,true);}
    public void DuplicateChart(object control){try{PowerPointBridge.DuplicateSelected(application);RefreshRibbon();}catch(Exception ex){MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
    public void FitChart(object control){try{PowerPointBridge.FitSelected(application);}catch(Exception ex){MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
    public void ShowDiagnostics(object control){using(SupportForm form=new SupportForm(InstallationHealth.Collect(typeof(Connect).Assembly.GetName().Version.ToString(3)))){try{form.ShowDialog(new PowerPointOwner(application));}catch{form.ShowDialog();}}}
    public void ShowHelp(object control){MessageBox.Show("Mello Charts "+typeof(Connect).Assembly.GetName().Version.ToString(3)+"\n\n1. Selecione um slide e escolha o gráfico nesta aba.\n2. Edite ou importe os dados e escolha o estilo no painel lateral. O editor permite continuar navegando no PowerPoint.\n3. Clique em Inserir no slide. O destino é o slide em que o editor foi aberto.\n4. Para editar depois, selecione o grupo ou um elemento dele e use Editar dados.\n5. Salve a apresentação e um backup JSON do projeto.\n\nDuplicar gráfico preserva ajustes nativos da cópia. Ajustar ao slide centraliza em 16:9 e remove a rotação.\n\nRecuperar rascunho abre o último estado válido preservado ao fechar a janela do PowerPoint.\n\nDiagnóstico apenas verifica a instalação. Exportar PNG salva uma imagem em 1920 × 1080.\n\nAtualizar recria o grupo e substitui edições manuais internas. Mantenha o grupo e os metadados para recuperar os dados. Use o instalador para remover o suplemento.","Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Information);}
    private void Open(string type,bool edit){
      try{
        if(application==null)throw new InvalidOperationException("O PowerPoint não conectou o suplemento. Feche e abra o PowerPoint novamente.");
        ChartReference reference=null;ChartDocument doc;
        if(edit){reference=PowerPointBridge.Selected(application);doc=PowerPointBridge.ReadDocument(application,reference);}else{PowerPointBridge.CurrentSlide(application);doc=ChartCore.Example(type);}
        if(editor!=null && !editor.IsDisposed){editor.Close();if(!editor.IsDisposed){editor.Activate();return;}}
        EditorForm next=new EditorForm(application,doc,reference);editor=next;next.OwnerWindowClosed+=(s,e)=>{if(Object.ReferenceEquals(editor,next))editor=null;RefreshRibbon();};next.FormClosed+=(s,e)=>{if(Object.ReferenceEquals(editor,next))editor=null;next.Dispose();RefreshRibbon();};try{next.Show(new PowerPointOwner(application));}catch{editor=null;next.Dispose();throw;}
      }catch(Exception ex){MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
    }
    private class PictureHost:AxHost {private PictureHost():base(""){}public static object Convert(Image image){return GetIPictureDispFromPicture(image);}}
  }

  public sealed class ChartReference {public int SlideId;public int ShapeId;public string ChartId;}
  public static class Metadata {
    public const string Prefix="MelloCharts:";
    private const string CountTag="MELLO_COUNT",VersionTag="MELLO_VERSION";
    public static IList<string> Encode(ChartDocument doc){
      string json=ChartCore.Serialize(doc);byte[] bytes=Encoding.UTF8.GetBytes(json);if(bytes.Length>25000)throw new InvalidOperationException("O projeto excede o tamanho de metadados permitido.");
      string hex=BitConverter.ToString(bytes).Replace("-","");List<string> chunks=new List<string>();for(int i=0;i<hex.Length;i+=200)chunks.Add(hex.Substring(i,Math.Min(200,hex.Length-i)));return chunks;
    }
    public static ChartDocument Decode(IList<string> chunks){
      if(chunks==null||chunks.Count<1||chunks.Count>250)throw new InvalidOperationException("Metadados do gráfico inválidos. Use seu backup JSON.");
      if(chunks.Any(s=>s==null||s.Length>200))throw new InvalidOperationException("Bloco de dados inválido.");string hex=String.Join("",chunks);if(hex.Length%2!=0)throw new InvalidOperationException("Dados do gráfico incompletos.");
      try{byte[] bytes=new byte[hex.Length/2];for(int i=0;i<bytes.Length;i++)bytes[i]=Convert.ToByte(hex.Substring(i*2,2),16);return ChartCore.Deserialize(new UTF8Encoding(false,true).GetString(bytes));}catch(Exception ex){throw new InvalidOperationException("Não foi possível recuperar os dados. Abra seu backup JSON. "+ex.Message);}
    }
    public static void Write(object shape,ChartDocument doc){dynamic s=shape;IList<string> chunks=Encode(doc);s.Tags.Add(VersionTag,"1");s.Tags.Add(CountTag,chunks.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));for(int i=0;i<chunks.Count;i++)s.Tags.Add("MELLO_DATA_"+i.ToString("D3"),chunks[i]);}
    public static ChartDocument Read(object shape){
      dynamic s=shape;if((string)s.Tags[VersionTag]!="1")throw new InvalidOperationException("Este grupo não tem dados compatíveis com o Mello Charts Desktop.");
      int count;if(!Int32.TryParse((string)s.Tags[CountTag],out count)||count<1||count>250)throw new InvalidOperationException("Dados do gráfico incompletos. Use o JSON de backup.");
      List<string> chunks=new List<string>();for(int i=0;i<count;i++)chunks.Add((string)s.Tags["MELLO_DATA_"+i.ToString("D3")]);return Decode(chunks);
    }
  }
  public static partial class PowerPointBridge {
    public static object CurrentSlide(object application){
      dynamic app=application;if(app==null||app.Presentations.Count==0)throw new InvalidOperationException("Abra uma apresentação e selecione um slide.");
      try{dynamic window=app.ActiveWindow;if(window==null)throw new InvalidOperationException();dynamic slide=window.View.Slide;if(slide==null)throw new InvalidOperationException();int id=slide.SlideID;return slide;}catch{throw new InvalidOperationException("Use a exibição Normal do PowerPoint e selecione um único slide.");}
    }
    public static ChartReference Selected(object application){
      dynamic app=application,slide=CurrentSlide(application);
      try{dynamic selection=app.ActiveWindow.Selection;int type=(int)selection.Type;if(type!=2 && type!=3)throw new InvalidOperationException();dynamic range=selection.ShapeRange;try{if(Convert.ToBoolean(selection.HasChildShapeRange))range=selection.ChildShapeRange;}catch{}if((int)range.Count<1)throw new InvalidOperationException();ChartReference result=null;
        for(int i=1;i<=(int)range.Count;i++){ChartReference next=AncestorChart((object)slide,(object)range[i]);if(result!=null && result.ShapeId!=next.ShapeId)throw new InvalidOperationException();result=next;}return result;
      }catch{throw new InvalidOperationException("Selecione um gráfico Mello Charts ou elementos de um único gráfico. Não misture gráficos diferentes na seleção.");}
    }
    private static object FindShape(object slide,ChartReference target){
      dynamic s=slide;if((int)s.SlideID!=target.SlideId)throw new InvalidOperationException("Volte ao slide do gráfico que estava editando.");
      for(int i=1;i<=(int)s.Shapes.Count;i++){dynamic sh=s.Shapes[i];if((int)sh.Id==target.ShapeId){if((string)sh.Name!=Metadata.Prefix+target.ChartId||(int)sh.Type!=6)throw new InvalidOperationException("O grupo original foi alterado. Feche o editor e selecione o gráfico novamente.");return sh;}}
      throw new InvalidOperationException("O gráfico original não está mais neste slide. Salve o JSON e insira um novo gráfico.");
    }
    public static ChartDocument ReadDocument(object app,ChartReference target){return Metadata.Read(FindShape(CurrentSlide(app),target));}
    private static int ColorValue(string hex){return ColorTranslator.ToOle(ColorTranslator.FromHtml(hex));}
    public static ChartReference InsertOrUpdate(object application,ChartDocument doc,ChartReference target){ChartCore.Validate(doc);object slide=CurrentSlide(application);UndoEntry(application);return DrawOnSlide(slide,doc,target);}
    public static ChartReference DrawOnSlide(object slideObject,ChartDocument doc,ChartReference target){
      return DrawOnSlide(slideObject,doc,target,null);
    }
    private static ChartReference DrawOnSlide(object slideObject,ChartDocument doc,ChartReference target,ChartBounds supplied){
      ChartCore.Validate(doc);Scene scene=ChartCore.BuildScene(doc);dynamic slide=slideObject;dynamic original=null;
      double left=24,top=24,width=640,height=360,rotation=0;
      if(target!=null){original=FindShape(slide,target);Metadata.Read(original);left=(double)original.Left;top=(double)original.Top;width=(double)original.Width;height=(double)original.Height;rotation=(double)original.Rotation;}
      else{ChartBounds fit=supplied??Placement(slideObject);left=fit.Left;top=fit.Top;width=fit.Width;height=fit.Height;rotation=fit.Rotation;}
      if(width<=0||height<=0)throw new InvalidOperationException("O tamanho do gráfico é inválido.");
      string id=Guid.NewGuid().ToString(),partPrefix="MelloPart:"+id+":";List<string> names=new List<string>();List<object> created=new List<object>();dynamic group=null;
      double sx=width/scene.Width,sy=height/scene.Height,fontScale=Math.Min(sx,sy);
      try{
        dynamic canvas=slide.Shapes.AddShape(1,(float)left,(float)top,(float)width,(float)height);created.Add((object)canvas);canvas.Name=partPrefix+"canvas";names.Add((string)canvas.Name);canvas.Fill.Solid();canvas.Fill.ForeColor.RGB=ColorValue("#FFFFFF");canvas.Line.Visible=0;
        foreach(SceneElement e in scene.Elements){
          dynamic shape;float x=(float)(left+e.X*sx),y=(float)(top+e.Y*sy),w=(float)(e.W*sx),h=(float)(e.H*sy);
          if(e.Type=="rect"||e.Type=="ellipse"){
            if(w<=0||h<=0)continue;shape=slide.Shapes.AddShape(e.Type=="ellipse"?9:1,x,y,w,h);created.Add((object)shape);shape.Name=partPrefix+names.Count;names.Add((string)shape.Name);shape.Fill.Solid();shape.Fill.ForeColor.RGB=ColorValue(e.Color);shape.Line.Visible=0;
          }else if(e.Type=="polygon"){
            dynamic freeform=slide.Shapes.BuildFreeform(1,(float)(left+e.Points[0]*sx),(float)(top+e.Points[1]*sy));
            // Straight segments require msoEditingAuto, even when the first node is a corner.
            for(int i=2;i<e.Points.Length;i+=2)freeform.AddNodes(0,0,(float)(left+e.Points[i]*sx),(float)(top+e.Points[i+1]*sy));
            shape=freeform.ConvertToShape();created.Add((object)shape);shape.Name=partPrefix+names.Count;names.Add((string)shape.Name);shape.Fill.Solid();shape.Fill.ForeColor.RGB=ColorValue(e.Color);shape.Line.Visible=0;
          }else if(e.Type=="line"){
            shape=slide.Shapes.AddLine(x,y,(float)(left+e.X2*sx),(float)(top+e.Y2*sy));created.Add((object)shape);shape.Name=partPrefix+names.Count;names.Add((string)shape.Name);shape.Line.ForeColor.RGB=ColorValue(e.Color);shape.Line.Weight=(float)Math.Max(.25,e.Size*fontScale);if(e.Dash)shape.Line.DashStyle=4;
          }else if(e.Type=="text"){
            shape=slide.Shapes.AddTextbox(1,x,y,w,h);created.Add((object)shape);shape.Name=partPrefix+names.Count;names.Add((string)shape.Name);shape.Fill.Visible=0;shape.Line.Visible=0;dynamic frame=shape.TextFrame;frame.MarginLeft=0;frame.MarginTop=0;frame.MarginRight=0;frame.MarginBottom=0;frame.AutoSize=0;frame.WordWrap=-1;
            dynamic range=frame.TextRange;range.Text=e.Text;range.Font.Name="Arial";range.Font.Size=(float)Math.Max(1,e.Size*fontScale);range.Font.Bold=e.Bold?-1:0;range.Font.Color.RGB=ColorValue(e.Color);range.ParagraphFormat.Alignment=e.Align=="center"?2:e.Align=="right"?3:1;
          }else throw new InvalidOperationException("Elemento de desenho desconhecido.");
        }
        group=slide.Shapes.Range(names.ToArray()).Group();group.Name=Metadata.Prefix+id;group.Rotation=(float)rotation;
        group.AlternativeText="Mello Charts — "+doc.title+". "+doc.type+". "+doc.rows.Count+" categorias. Use Editar dados para recuperar os dados.";
        Metadata.Write(group,doc);ChartDocument verified=Metadata.Read(group);if(ChartCore.Serialize(verified)!=ChartCore.Serialize(doc))throw new InvalidOperationException("Os dados não foram confirmados no grupo.");
      }catch(Exception ex){
        try{if(group!=null)group.Delete();else for(int i=created.Count-1;i>=0;i--){dynamic s=created[i];s.Delete();}}
        catch(Exception cleanup){throw new InvalidOperationException("Falha ao criar o gráfico: "+ex.Message+". Confira possíveis formas novas no slide: "+cleanup.Message);}
        throw new InvalidOperationException((target!=null?"O gráfico anterior foi preservado. ":"")+"Não foi possível criar o novo gráfico: "+ex.Message);
      }
      if(original!=null){try{dynamic checkedOriginal=FindShape(slide,target);checkedOriginal.Delete();}catch(Exception ex){throw new InvalidOperationException("O novo gráfico está pronto, mas o anterior não pôde ser removido. Confira os dois grupos antes de excluir um. "+ex.Message);}}
      ChartReference next=new ChartReference{SlideId=(int)slide.SlideID,ShapeId=(int)group.Id,ChartId=id};try{group.Select();}catch{/* Insertion succeeds even if selection isn't available. */}return next;
    }
  }
}
