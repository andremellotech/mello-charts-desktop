using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace MelloCharts {
  public sealed class ChartContext {internal object Presentation;public int SlideId,SlideNumber;}
  public sealed class ChartBounds {public double Left,Top,Width,Height,Rotation;}
  public static partial class PowerPointBridge {
    static bool SameIdentity(object a,object b){
      if(Object.ReferenceEquals(a,b))return true;if(a==null || b==null || !Marshal.IsComObject(a) || !Marshal.IsComObject(b))return false;
      IntPtr first=IntPtr.Zero,second=IntPtr.Zero;try{first=Marshal.GetIUnknownForObject(a);second=Marshal.GetIUnknownForObject(b);return first==second;}finally{if(first!=IntPtr.Zero)Marshal.Release(first);if(second!=IntPtr.Zero)Marshal.Release(second);}
    }
    public static ChartContext CaptureContext(object application){dynamic app=application,slide=CurrentSlide(application);ChartContext context=new ChartContext{Presentation=(object)app.ActivePresentation,SlideId=(int)slide.SlideID};try{context.SlideNumber=(int)slide.SlideIndex;}catch{}return context;}
    public static object ResolveContext(object application,ChartContext context){
      if(context==null)throw new InvalidOperationException("Abra o editor novamente pela aba Mello Charts.");
      dynamic app=application;bool open=false;try{for(int i=1;i<=(int)app.Presentations.Count;i++)if(SameIdentity((object)app.Presentations[i],context.Presentation)){open=true;break;}}catch{}
      if(!open)throw new InvalidOperationException("A apresentação de origem foi fechada. Salve o projeto JSON e abra um novo editor.");
      if(!SameIdentity((object)app.ActivePresentation,context.Presentation))throw new InvalidOperationException("Volte à apresentação em que este editor foi aberto. O gráfico não será inserido em outro arquivo.");
      CurrentSlide(application);
      try{dynamic presentation=context.Presentation;return (object)presentation.Slides.FindBySlideID(context.SlideId);}catch{throw new InvalidOperationException("O slide de origem foi excluído. Salve o JSON e abra um novo gráfico em outro slide.");}
    }
    public static ChartBounds FitBounds(double slideWidth,double slideHeight){
      if(Double.IsNaN(slideWidth)||Double.IsInfinity(slideWidth)||Double.IsNaN(slideHeight)||Double.IsInfinity(slideHeight)||slideWidth<=0||slideHeight<=0)throw new ArgumentException("Dimensões do slide inválidas.");
      double margin=Math.Min(24,Math.Min(slideWidth,slideHeight)*.08),width=Math.Min(slideWidth-2*margin,(slideHeight-2*margin)*16/9),height=width*9/16;
      return new ChartBounds{Left=(slideWidth-width)/2,Top=(slideHeight-height)/2,Width=width,Height=height};
    }
    static double[] SlideSize(object slideObject){
      dynamic slide=slideObject;try{dynamic parent=slide.Parent;try{return new[]{(double)parent.PageSetup.SlideWidth,(double)parent.PageSetup.SlideHeight};}catch{return new[]{(double)parent.Parent.PageSetup.SlideWidth,(double)parent.Parent.PageSetup.SlideHeight};}}catch{throw new InvalidOperationException("Não foi possível ler o tamanho do slide. Volte à exibição Normal e tente novamente.");}
    }
    static ChartBounds Placement(object slide){double[] page=SlideSize(slide);return FitBounds(page[0],page[1]);}
    static void UndoEntry(object application){dynamic app=application;app.StartNewUndoEntry();}
    public static ChartReference InsertInContext(object app,ChartDocument document,ChartReference target,ChartContext context){object slide=ResolveContext(app,context);if(target!=null && target.SlideId!=context.SlideId)throw new InvalidOperationException("O gráfico pertence a outro slide.");ChartCore.Validate(document);UndoEntry(app);return DrawOnSlide(slide,document,target);}
    static ChartReference ReferenceFor(object slideObject,object shapeObject){dynamic slide=slideObject,shape=shapeObject;Guid id;string name=(string)shape.Name;if((int)shape.Type!=6||!name.StartsWith(Metadata.Prefix,StringComparison.Ordinal)||!Guid.TryParse(name.Substring(Metadata.Prefix.Length),out id))throw new InvalidOperationException("Selecione um gráfico Mello Charts ou um de seus elementos.");return new ChartReference{SlideId=(int)slide.SlideID,ShapeId=(int)shape.Id,ChartId=id.ToString()};}
    static ChartReference AncestorChart(object slide,object shapeObject){dynamic shape=shapeObject;for(int depth=0;depth<16 && shape!=null;depth++){try{return ReferenceFor(slide,(object)shape);}catch{}try{shape=shape.ParentGroup;}catch{break;}}throw new InvalidOperationException("Este elemento não pertence a um gráfico Mello Charts.");}
    public static ChartReference DuplicateSelected(object application){
      dynamic slide=CurrentSlide(application);ChartReference reference=Selected(application);dynamic original=FindShape(slide,reference);ChartDocument document=Metadata.Read(original);double[] page=SlideSize(slide);
      ChartBounds bounds=new ChartBounds{Left=(double)original.Left+18,Top=(double)original.Top+18,Width=(double)original.Width,Height=(double)original.Height,Rotation=(double)original.Rotation};
      if(bounds.Width>page[0] || bounds.Height>page[1]){double factor=Math.Min(page[0]/bounds.Width,page[1]/bounds.Height)*.92;bounds.Width*=factor;bounds.Height*=factor;}
      bounds.Left=Math.Max(0,Math.Min(bounds.Left,page[0]-bounds.Width));bounds.Top=Math.Max(0,Math.Min(bounds.Top,page[1]-bounds.Height));UndoEntry(application);
      dynamic copy=null;
      try{
        dynamic range=original.Duplicate();if((int)range.Count!=1)throw new InvalidOperationException("O PowerPoint retornou uma cópia inesperada.");copy=range[1];
        string id=Guid.NewGuid().ToString();copy.Name=Metadata.Prefix+id;int locked=(int)copy.LockAspectRatio;copy.LockAspectRatio=0;copy.Width=(float)bounds.Width;copy.Height=(float)bounds.Height;copy.Left=(float)bounds.Left;copy.Top=(float)bounds.Top;copy.LockAspectRatio=locked;
        Metadata.Write(copy,document);if(ChartCore.Serialize(Metadata.Read(copy))!=ChartCore.Serialize(document))throw new InvalidOperationException("Os dados da cópia não foram confirmados.");
        ChartReference result=new ChartReference{SlideId=(int)slide.SlideID,ShapeId=(int)copy.Id,ChartId=id};try{copy.Select();}catch{}return result;
      }catch(Exception ex){if(copy!=null)try{copy.Delete();}catch{throw new InvalidOperationException("A duplicação falhou. O original foi preservado, mas confira possíveis cópias parciais no slide. "+ex.Message);}throw new InvalidOperationException("A duplicação foi cancelada; o original foi preservado. "+ex.Message);}
    }
    public static void FitSelected(object application){
      dynamic slide=CurrentSlide(application);ChartReference reference=Selected(application);dynamic shape=FindShape(slide,reference);Metadata.Read(shape);ChartBounds fit=Placement((object)slide);
      double left=(double)shape.Left,top=(double)shape.Top,width=(double)shape.Width,height=(double)shape.Height,rotation=(double)shape.Rotation;int locked=(int)shape.LockAspectRatio;UndoEntry(application);
      try{shape.LockAspectRatio=0;shape.Rotation=0;shape.Width=(float)fit.Width;shape.Height=(float)fit.Height;shape.Left=(float)fit.Left;shape.Top=(float)fit.Top;shape.LockAspectRatio=-1;shape.Select();}
      catch(Exception ex){try{shape.LockAspectRatio=0;shape.Width=(float)width;shape.Height=(float)height;shape.Left=(float)left;shape.Top=(float)top;shape.Rotation=(float)rotation;shape.LockAspectRatio=locked;}catch{throw new InvalidOperationException("O ajuste falhou e não foi possível restaurar toda a geometria. Confira o gráfico e use Desfazer. "+ex.Message);}throw new InvalidOperationException("O ajuste foi cancelado e a geometria anterior foi restaurada. "+ex.Message);}
    }
  }
  public static class RecoveryStore {
    public static string Folder{get{return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"MelloCharts","Drafts");}}
    public static string SaveTo(string folder,ChartDocument document){ChartCore.Validate(document);Directory.CreateDirectory(folder);string path=Path.Combine(folder,"rascunho-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")+".json");using(FileStream stream=new FileStream(path,FileMode.CreateNew,FileAccess.Write,FileShare.None))using(StreamWriter writer=new StreamWriter(stream,new UTF8Encoding(false)))writer.Write(ChartCore.Serialize(document));return path;}
  }
  internal sealed class PowerPointOwner:IWin32Window {public IntPtr Handle{get;private set;}public PowerPointOwner(object app){dynamic host=app;Handle=new IntPtr((long)host.ActiveWindow.HWND);if(Handle==IntPtr.Zero)throw new InvalidOperationException("A janela do PowerPoint não está disponível.");}}
  public static class RibbonIcons {
    public static Bitmap Create(string id){
      Bitmap image=new Bitmap(32,32);using(Graphics g=Graphics.FromImage(image))using(Pen navy=new Pen(Color.FromArgb(36,79,115),2.3f))using(Brush blue=new SolidBrush(Color.FromArgb(36,79,115)))using(Brush green=new SolidBrush(Color.FromArgb(97,169,155)))using(Brush amber=new SolidBrush(Color.FromArgb(189,128,80))){
        g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        if(id.Contains("Gallery")){g.FillRectangle(blue,3,3,11,11);g.FillRectangle(green,18,3,11,11);g.FillRectangle(green,3,18,11,11);g.FillRectangle(amber,18,18,11,11);}
        else if(id.Contains("Build")){for(int i=0;i<3;i++){int x=4+i*9;g.FillRectangle(blue,x,19-i*3,6,10+i*3);g.FillRectangle(green,x,11-i*3,6,8);}}
        else if(id.Contains("Waterfall")){g.FillRectangle(blue,3,12,6,17);g.FillRectangle(amber,13,5,6,7);g.FillRectangle(green,23,5,6,5);g.DrawLine(navy,9,12,13,12);g.DrawLine(navy,19,5,23,5);}
        else if(id.Contains("Columns")){g.FillRectangle(blue,3,14,5,15);g.FillRectangle(green,9,9,5,20);g.FillRectangle(blue,19,5,5,24);g.FillRectangle(green,25,12,5,17);}
        else if(id.Contains("Duplicate")){g.DrawRectangle(navy,4,3,17,20);g.FillRectangle(Brushes.White,11,10,17,20);g.DrawRectangle(navy,11,10,17,20);}
        else if(id.Contains("Fit")){g.DrawRectangle(navy,3,5,26,22);g.DrawLine(navy,8,21,23,11);g.DrawLines(navy,new[]{new Point(16,11),new Point(23,11),new Point(23,18)});}
        else if(id.Contains("Edit")){g.DrawRectangle(navy,4,7,21,22);g.FillPolygon(green,new[]{new Point(11,24),new Point(12,18),new Point(25,4),new Point(29,8),new Point(16,23)});}
        else if(id.Contains("Diagnostics")){g.DrawEllipse(navy,3,3,18,18);g.DrawLine(navy,19,19,28,29);g.DrawLines(navy,new[]{new Point(7,12),new Point(10,16),new Point(17,8)});}
        else{g.DrawEllipse(navy,3,3,26,26);using(Font font=new Font("Arial",20,FontStyle.Bold,GraphicsUnit.Pixel))g.DrawString("?",font,blue,9,4);}
      }return image;
    }
  }
}
