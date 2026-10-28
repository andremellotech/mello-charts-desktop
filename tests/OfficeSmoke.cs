using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using MelloCharts;
public static class OfficeSmoke {
  [STAThread]public static int Main(string[] args){
    if(args.Length!=1){Console.Error.WriteLine("Informe uma pasta existente para os PPTX de teste. Não instala o suplemento.");return 2;}
    string folder=Path.GetFullPath(args[0]);if(!Directory.Exists(folder))return 2;
    Process[] processes=Process.GetProcessesByName("POWERPNT");bool alreadyRunning=processes.Length>0;foreach(Process process in processes)process.Dispose();
    // Keep cleanup arguments statically typed. Dynamic dispatch can query a dead
    // COM server even for Marshal.IsComObject after Quit (RPC_E_INVALID_OBJECT).
    object application=null,presentation=null,presentations=null,slides=null;
    try{
      Type type=Type.GetTypeFromProgID("PowerPoint.Application",true);application=Activator.CreateInstance(type);dynamic app=application;presentations=(object)app.Presentations;dynamic decks=presentations;presentation=(object)decks.Add(0);dynamic pres=presentation;slides=(object)pres.Slides;dynamic slideCollection=slides;
      foreach(ChartDefinition definition in ChartCatalog.All){
        string chartType=definition.Key;object slideObject=null,shapesObject=null,shapeObject=null,updatedObject=null;
        try{
          slideObject=(object)slideCollection.Add((int)slideCollection.Count+1,12);dynamic slide=slideObject;ChartDocument d=ChartCore.Example(chartType);ChartReference first=PowerPointBridge.DrawOnSlide(slideObject,d,null);
          shapesObject=(object)slide.Shapes;dynamic shapes=shapesObject;shapeObject=(object)shapes[1];dynamic shape=shapeObject;shape.Left=62;shape.Top=48;shape.Width=560;shape.Height=315;shape.Rotation=7;
          double expectedLeft=(double)shape.Left,expectedTop=(double)shape.Top,expectedWidth=(double)shape.Width,expectedHeight=(double)shape.Height,expectedRotation=(double)shape.Rotation;
          Console.WriteLine("BEFORE "+chartType+": left="+expectedLeft+", top="+expectedTop+", width="+expectedWidth+", height="+expectedHeight+", rotation="+expectedRotation);
          Release(shapeObject);shapeObject=null;
          d.title="Validado no PowerPoint · "+chartType;ChartReference next=PowerPointBridge.DrawOnSlide(slideObject,d,first);
          if((int)shapes.Count!=1)throw new Exception("Contagem de grupos inválida");updatedObject=(object)shapes[1];dynamic updated=updatedObject;
          double actualLeft=(double)updated.Left,actualTop=(double)updated.Top,actualWidth=(double)updated.Width,actualHeight=(double)updated.Height,actualRotation=(double)updated.Rotation;
          Console.WriteLine("AFTER "+chartType+": left="+actualLeft+", top="+actualTop+", width="+actualWidth+", height="+actualHeight+", rotation="+actualRotation);
          if(Math.Abs(actualLeft-expectedLeft)>.2||Math.Abs(actualTop-expectedTop)>.2||Math.Abs(actualWidth-expectedWidth)>.2||Math.Abs(actualHeight-expectedHeight)>.2||Math.Abs(actualRotation-expectedRotation)>.2)throw new Exception("Geometria não preservada em "+chartType);
          ChartDocument saved=Metadata.Read(updatedObject);if(ChartCore.Serialize(saved)!=ChartCore.Serialize(d))throw new Exception("Dados ou estilo não recuperados");Console.WriteLine("OK "+chartType+": grupo, metadados e atualização confirmados.");
        }finally{Release(updatedObject);Release(shapeObject);Release(shapesObject);Release(slideObject);}
      }
      string output=Path.Combine(folder,"mello-charts-office-smoke-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".pptx");pres.SaveAs(output,24);Console.WriteLine(output);return 0;
    }catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
    finally{Release(slides);Close(presentation);Release(presentation);Release(presentations);if(!alreadyRunning)Quit(application);Release(application);}
  }
  static void Close(object presentation){if(presentation==null)return;try{dynamic deck=presentation;deck.Saved=-1;deck.Close();}catch(COMException){}catch(InvalidComObjectException){}}
  static void Quit(object application){if(application==null)return;try{dynamic app=application;app.Quit();}catch(COMException){}catch(InvalidComObjectException){}}
  static void Release(object value){if(value==null)return;try{if(Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}catch(COMException){}catch(InvalidComObjectException){}}
}
