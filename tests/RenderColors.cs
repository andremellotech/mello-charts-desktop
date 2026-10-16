using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using MelloCharts;

// Render the application's own controls and a fictional custom-color example.
public static class RenderColors {
    [STAThread] public static void Main(string[] args) {
        string output=args[0];Directory.CreateDirectory(output);
        ChartDocument d=ChartCore.Example("buildup");
        d.title="A identidade da sua marca em cada gráfico";
        d.subtitle="Cores RGB escolhidas pelo usuário, com valores editáveis e contraste automático";
        d.customColors=new ChartColors {actual="#6D28D9",forecast="#FBBF24",baseline="#E11D48"};
        File.WriteAllText(Path.Combine(output,"cores-personalizadas.json"),ChartCore.Serialize(d),new UTF8Encoding(false));
        File.WriteAllText(Path.Combine(output,"cores-personalizadas.csv"),ChartCore.ToCsv(d),new UTF8Encoding(false));
        using(Bitmap image=ChartPreview.RenderImage(ChartCore.BuildScene(d),1920,1080))
            image.Save(Path.Combine(output,"cores-personalizadas.png"),System.Drawing.Imaging.ImageFormat.Png);
        using(ChartColorForm form=new ChartColorForm(d))Render(form,Path.Combine(output,"editor-de-cores.png"));
        File.AppendAllText(Path.Combine(output,"README.md"),"\n## Cores personalizadas\n\nUse **Cores personalizadas…** no painel do editor. Digite HEX/RGB ou abra o seletor; configure séries, fundo e fatias e restaure o padrão quando precisar.\n\n[Projeto JSON](cores-personalizadas.json) · [Tabela CSV](cores-personalizadas.csv)\n\n![Exemplo com cores personalizadas](cores-personalizadas.png)\n\n![Janela de cores](editor-de-cores.png)\n",new UTF8Encoding(false));
        Console.WriteLine("cores-personalizadas");
    }
    static void Render(Form form,string path) {
        form.ShowInTaskbar=false;form.Opacity=0;form.Show();Application.DoEvents();
        using(Bitmap image=new Bitmap(form.Width,form.Height)){form.DrawToBitmap(image,new Rectangle(0,0,image.Width,image.Height));image.Save(path,System.Drawing.Imaging.ImageFormat.Png);}
    }
}
