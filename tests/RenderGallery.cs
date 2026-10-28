using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Text;
using MelloCharts;

// Uses the same drawing code as the editor's preview and PNG export. No Office activation.
public static class RenderGallery {
  [STAThread] public static void Main(string[] args){
    string output=args.Length>0?args[0]:"examples/gallery";Directory.CreateDirectory(output);
    string[] types={"buildup","waterfall","columns","bars","lines","area","combo","stacked100","pie","donut","pareto","gantt","scatter","bubble"};
    string[] styles={"executive","minimal","dark","minimal","executive","executive","dark","minimal","executive","dark","executive","executive","minimal","dark"};
    string[] names={"buildup-executivo","cascata-essencial","comparativo-escuro","barras-horizontais","linhas","area-comparativa","combinado","empilhado-100","pizza","rosca","pareto","gantt","dispersao","bolhas"};
    StringBuilder readme=new StringBuilder("# Galeria de gráficos\n\nTodos os dados são fictícios. Abra um JSON pelo botão **Abrir projeto**, ou importe o CSV no tipo correspondente. Os PNGs usam o mesmo desenho da prévia e da exportação.\n\n");
    using(Bitmap sheet=new Bitmap(1440,1620))using(Graphics gallery=Graphics.FromImage(sheet))
    using(Font heading=new Font("Arial",25,FontStyle.Bold,GraphicsUnit.Pixel))
    using(Font caption=new Font("Arial",15,FontStyle.Bold,GraphicsUnit.Pixel))
    using(Brush ink=new SolidBrush(Color.FromArgb(32,55,77))) {
    gallery.Clear(Color.FromArgb(239,243,246));gallery.DrawString("Mello Charts · 14 maneiras de apresentar seus dados",heading,ink,28,20);
    for(int i=0;i<types.Length;i++){
      ChartDocument document=ChartCore.Example(types[i]);document.style=styles[i];document.palette=i==1?"forest":"ocean";document.showGrid=i!=1;
      if(i==0){document.title="Investimento por frente de trabalho";document.subtitle="Realizado e saldo projetado compõem o EAC. O traço indica o baseline.";}
      if(i==2){document.title="Execução e projeção por período";document.subtitle="Comparação das séries com a referência aprovada";}
      File.WriteAllText(Path.Combine(output,names[i]+".json"),ChartCore.Serialize(document),new UTF8Encoding(false));
      File.WriteAllText(Path.Combine(output,names[i]+".csv"),ChartCore.ToCsv(document),new UTF8Encoding(false));
      using(Bitmap bitmap=ChartPreview.RenderImage(ChartCore.BuildScene(document),1920,1080)) {
        bitmap.Save(Path.Combine(output,names[i]+".png"),System.Drawing.Imaging.ImageFormat.Png);
        int x=24+(i%3)*472,y=72+(i/3)*302;
        gallery.DrawImage(bitmap,new Rectangle(x,y,448,252));
        gallery.DrawString(ChartCatalog.Get(types[i]).Name,caption,ink,x+4,y+264);
      }
      readme.Append("## ").Append(ChartCatalog.Get(types[i]).Name).Append("\n\n").Append(ChartCatalog.Get(types[i]).Description).Append("\n\n[Projeto JSON](").Append(names[i]).Append(".json) · [Tabela CSV](").Append(names[i]).Append(".csv)\n\n![").Append(ChartCatalog.Get(types[i]).Name).Append("](").Append(names[i]).Append(".png)\n\n");
      Console.WriteLine(names[i]);
    }
    sheet.Save(Path.Combine(output,"catalogo.png"),System.Drawing.Imaging.ImageFormat.Png);
    }
    File.WriteAllText(Path.Combine(output,"README.md"),readme.ToString().TrimEnd()+"\n",new UTF8Encoding(false));
  }
}
