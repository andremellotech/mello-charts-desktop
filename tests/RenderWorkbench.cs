using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using MelloCharts;

public static class RenderWorkbench {
    [STAThread] public static int Main(string[] args) {
        Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
        string output=args.Length>0?args[0]:"examples/gallery";Directory.CreateDirectory(output);
        using(EditorForm editor=new EditorForm(null,ChartCore.Example("gantt"),null)) {
            editor.ShowInTaskbar=false;editor.Opacity=0;editor.Show();Application.DoEvents();
            using(Bitmap image=new Bitmap(editor.Width,editor.Height)){editor.DrawToBitmap(image,new Rectangle(Point.Empty,editor.Size));image.Save(Path.Combine(output,"editor-gantt.png"));}
        }
        using(ChartGalleryForm gallery=new ChartGalleryForm()) {
            gallery.ShowInTaskbar=false;gallery.Opacity=0;gallery.Show();Application.DoEvents();
            using(Bitmap image=new Bitmap(gallery.Width,gallery.Height)){gallery.DrawToBitmap(image,new Rectangle(Point.Empty,gallery.Size));image.Save(Path.Combine(output,"galeria-modelos.png"));}
        }
        return 0;
    }
}
