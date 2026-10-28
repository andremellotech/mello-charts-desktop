using System;
using System.IO;
using System.Windows.Forms;
using MelloCharts;

internal static class DesktopLauncher {
    [STAThread] private static int Main(string[] args) {
        Application.EnableVisualStyles();Application.SetCompatibleTextRenderingDefault(false);
        try {
            ChartDocument document;
            if(args.Length>1)throw new ArgumentException("Abra um único projeto JSON por vez.");
            if(args.Length==1) {
                if(new FileInfo(args[0]).Length>1000000)throw new ArgumentException("O projeto deve ter até 1 MB.");
                document=ChartCore.Deserialize(File.ReadAllText(args[0],new System.Text.UTF8Encoding(false,true)));
            } else {
                using(ChartGalleryForm gallery=new ChartGalleryForm()) {
                    gallery.ShowInTaskbar=true;if(gallery.ShowDialog()!=DialogResult.OK)return 0;
                    document=ChartCore.Example(gallery.SelectedType);
                }
            }
            Application.Run(new EditorForm(null,document,null));return 0;
        } catch(Exception ex) {MessageBox.Show(ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);return 1;}
    }
}
