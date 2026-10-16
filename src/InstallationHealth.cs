using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MelloCharts {
  // Read-only diagnostics. Never activates Office or changes registry/security settings.
  public static class InstallationHealth {
    public static string Collect(string version){
      StringBuilder report=new StringBuilder("Mello Charts "+version+"\r\nDiagnóstico local de instalação\r\n\r\n");
      report.AppendLine("Windows: "+(Environment.Is64BitOperatingSystem?"64":"32")+" bits");report.AppendLine("Processo atual: "+(Environment.Is64BitProcess?"64":"32")+" bits");report.AppendLine("Runtime .NET: "+Environment.Version);
      string dll=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"MelloCharts","MelloCharts.AddIn.dll");
      try{report.AppendLine("DLL instalada: "+(File.Exists(dll)?AssemblyName.GetAssemblyName(dll).Version.ToString():"não encontrada"));}catch(Exception ex){report.AppendLine("DLL instalada: não foi possível verificar ("+ex.GetType().Name+")");}
      foreach(RegistryView view in Environment.Is64BitOperatingSystem?new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32}){
        report.AppendLine("\r\nRegistro HKCU / "+view+":");
        try{using(RegistryKey user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
          using(RegistryKey addin=user.OpenSubKey(@"Software\Microsoft\Office\PowerPoint\Addins\MelloCharts.Connect",false))report.AppendLine("  Suplemento / LoadBehavior: "+(addin==null?"não registrado":Convert.ToString(addin.GetValue("LoadBehavior","ausente"))));
          using(RegistryKey com=user.OpenSubKey(@"Software\Classes\CLSID\{49607523-6CCF-4559-97E4-FCF5BE438C3A}\InprocServer32",false)){
            report.AppendLine("  Classe COM: "+(com==null?"não registrada":"presente"));
            if(com!=null){report.AppendLine("  Assembly: "+Convert.ToString(com.GetValue("Assembly","ausente")));string uri=Convert.ToString(com.GetValue("CodeBase",""));Uri codeBase;report.AppendLine("  CodeBase local existente: "+(Uri.TryCreate(uri,UriKind.Absolute,out codeBase) && codeBase.IsFile && File.Exists(codeBase.LocalPath)?"sim":"não"));}
          }
        }}catch(Exception ex){report.AppendLine("  Leitura indisponível: "+ex.GetType().Name);}
      }
      report.AppendLine("\r\nEste relatório não comprova a ativação no PowerPoint. O teste final é abrir o aplicativo, localizar a aba e inserir um gráfico.");
      report.AppendLine("Se faltar o registro ou a DLL, feche o PowerPoint e use Instalar novamente. Se a organização bloquear suplementos, consulte a TI.");
      report.AppendLine("Os dados e os nomes das apresentações não são coletados. Nenhuma configuração foi alterada.");return report.ToString();
    }
  }
  public sealed class SupportForm:Form {
    public SupportForm(string report){
      Text="Mello Charts · Diagnóstico";Size=new Size(740,560);MinimumSize=new Size(540,380);StartPosition=FormStartPosition.CenterParent;Font=new Font("Segoe UI",9);Padding=new Padding(16);
      TextBox text=new TextBox{Text=report,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Vertical,Dock=DockStyle.Fill,BackColor=Color.White};Controls.Add(text);
      FlowLayoutPanel actions=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=42,FlowDirection=FlowDirection.RightToLeft};Button close=new Button{Text="Fechar",AutoSize=true},save=new Button{Text="Salvar relatório",AutoSize=true};close.Click+=(s,e)=>Close();save.Click+=(s,e)=>{using(SaveFileDialog dialog=new SaveFileDialog{Filter="Relatório de texto|*.txt",FileName="mello-charts-diagnostico.txt"}){if(dialog.ShowDialog(this)==DialogResult.OK)try{File.WriteAllText(dialog.FileName,report,new UTF8Encoding(true));}catch(Exception ex){MessageBox.Show(this,ex.Message,"Não foi possível salvar");}}};actions.Controls.Add(close);actions.Controls.Add(save);Controls.Add(actions);
    }
  }
}
