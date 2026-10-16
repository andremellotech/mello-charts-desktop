using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using MelloCharts;
public static class OfficeSmoke {
  [STAThread]public static int Main(string[] args){
    if(args.Length!=1){Console.Error.WriteLine("Informe uma pasta existente para os PPTX de teste. Não instala o suplemento.");return 2;}
    string folder=Path.GetFullPath(args[0]);if(!Directory.Exists(folder))return 2;bool alreadyRunning=Process.GetProcessesByName("POWERPNT").Length>0;dynamic app=null,pres=null;
    try{
      Type type=Type.GetTypeFromProgID("PowerPoint.Application",true);app=Activator.CreateInstance(type);pres=app.Presentations.Add(0);
      foreach(string chartType in new[]{"buildup","waterfall","columns"}){
        dynamic slide=pres.Slides.Add(pres.Slides.Count+1,12);ChartDocument d=ChartCore.Example(chartType);ChartReference first=PowerPointBridge.DrawOnSlide(slide,d,null);
        dynamic shape=slide.Shapes[1];shape.Left=62;shape.Top=48;shape.Width=560;shape.Height=315;shape.Rotation=7;
        d.title="Validado no PowerPoint · "+chartType;ChartReference next=PowerPointBridge.DrawOnSlide(slide,d,first);
        if(slide.Shapes.Count!=1)throw new Exception("Contagem de grupos inválida");dynamic updated=slide.Shapes[1];
        if(Math.Abs((double)updated.Left-62)>.2||Math.Abs((double)updated.Width-560)>.2||Math.Abs((double)updated.Rotation-7)>.2)throw new Exception("Geometria não preservada");
        ChartDocument saved=Metadata.Read(updated);if(saved.title!=d.title)throw new Exception("Título não recuperado");Console.WriteLine("OK "+chartType+": grupo, metadados e atualização confirmados.");
      }
      string output=Path.Combine(folder,"mello-charts-office-smoke-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".pptx");pres.SaveAs(output,24);Console.WriteLine(output);return 0;
    }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    finally{try{if(pres!=null){pres.Saved=-1;pres.Close();}}catch{}try{if(app!=null&&!alreadyRunning)app.Quit();}catch{}if(pres!=null&&Marshal.IsComObject(pres))Marshal.ReleaseComObject(pres);if(app!=null&&Marshal.IsComObject(app))Marshal.ReleaseComObject(app);}
  }
}
