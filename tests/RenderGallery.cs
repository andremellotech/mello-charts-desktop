using System;
using System.Drawing;
using System.IO;
using System.Text;
using MelloCharts;

// Uses the same drawing code as the editor's preview and PNG export. No Office activation.
public static class RenderGallery {
  [STAThread] public static void Main(string[] args){
    string output=args.Length>0?args[0]:"examples/gallery";Directory.CreateDirectory(output);
    string[] types={"buildup","waterfall","columns"},styles={"executive","minimal","dark"},names={"buildup-executivo","cascata-essencial","comparativo-escuro"};
    for(int i=0;i<types.Length;i++){
      ChartDocument document=ChartCore.Example(types[i]);document.style=styles[i];document.palette=i==1?"forest":"ocean";document.showGrid=i!=1;
      if(i==0){document.title="Investimento por frente de trabalho";document.subtitle="Realizado e saldo projetado compõem o EAC. O traço indica o baseline.";}
      if(i==2){document.title="Execução e projeção por período";document.subtitle="Comparação das séries com a referência aprovada";}
      File.WriteAllText(Path.Combine(output,names[i]+".json"),ChartCore.Serialize(document),new UTF8Encoding(false));
      using(Bitmap bitmap=ChartPreview.RenderImage(ChartCore.BuildScene(document),1920,1080))bitmap.Save(Path.Combine(output,names[i]+".png"),System.Drawing.Imaging.ImageFormat.Png);
      Console.WriteLine(names[i]);
    }
  }
}
