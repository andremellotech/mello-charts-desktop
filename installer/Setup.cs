// Original per-user installer for Mello Charts Desktop. No add-in code is loaded here.
// Build with the .NET Framework C# compiler and the MelloCharts.AddIn.dll resource.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Mello Charts Desktop Setup")]
[assembly: AssemblyDescription("Instalador por usuário do Mello Charts Desktop")]
[assembly: AssemblyCompany("Mello Charts")]
[assembly: AssemblyProduct("Mello Charts Desktop")]
[assembly: AssemblyVersion("0.5.0.0")]
[assembly: AssemblyFileVersion("0.5.0.0")]
[assembly: ComVisible(false)]

namespace MelloCharts.Installer
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                if (Environment.OSVersion.Platform != PlatformID.Win32NT)
                    throw new InvalidOperationException("Este instalador requer Windows.");

                bool uninstall = false;
                foreach (string arg in args)
                {
                    if (String.Equals(arg, "--uninstall", StringComparison.OrdinalIgnoreCase))
                        uninstall = true;
                }

                string sid;
                using (WindowsIdentity identity = WindowsIdentity.GetCurrent())
                    sid = identity.User.Value;
                bool first;
                using (Mutex mutex = new Mutex(true, "Local\\MelloCharts.Desktop.Setup." + sid, out first))
                {
                    if (!first)
                    {
                        MessageBox.Show("O instalador do Mello Charts já está aberto para este usuário.",
                            SetupConstants.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    try { Application.Run(new SetupWindow(uninstall)); }
                    finally { mutex.ReleaseMutex(); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Não foi possível abrir o instalador.\r\n\r\n" + ex.Message,
                    SetupConstants.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    internal static class SetupConstants
    {
        internal const string ProductName = "Mello Charts Desktop";
        internal const string DisplayVersion = "0.5.0";
        internal const string AssemblyVersion = "0.5.0.0";
        internal const string ProgId = "MelloCharts.Connect";
        internal const string ClassId = "{49607523-6CCF-4559-97E4-FCF5BE438C3A}";
        internal const string ManagedCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
        internal const string ResourceName = "MelloCharts.AddIn.dll";
        internal const string DllName = "MelloCharts.AddIn.dll";
        internal const string SetupName = "MelloCharts.Setup.exe";
        internal const string ClassKey = "Software\\Classes\\CLSID\\" + ClassId;
        internal const string ProgIdKey = "Software\\Classes\\" + ProgId;
        internal const string AddInKey = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\" + ProgId;
        internal const string UninstallKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\MelloChartsDesktop";
        internal static readonly string[] OwnKeys = { ClassKey, ProgIdKey, AddInKey, UninstallKey };

        internal static RegistryView[] RegistryViews()
        {
            return Environment.Is64BitOperatingSystem
                ? new RegistryView[] { RegistryView.Registry32, RegistryView.Registry64 }
                : new RegistryView[] { RegistryView.Registry32 };
        }

        internal static void RequirePowerPointClosed()
        {
            Process[] processes;
            try { processes = Process.GetProcessesByName("POWERPNT"); }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Não foi possível verificar se o PowerPoint está fechado. " + ex.Message);
            }
            bool open = processes.Length != 0;
            foreach (Process process in processes) process.Dispose();
            if (open)
                throw new InvalidOperationException("Feche todas as janelas do PowerPoint e tente novamente. " +
                    "O instalador não encerra o PowerPoint automaticamente.");
        }

        internal static void ValidateRegistryPath(string path)
        {
            foreach (string own in OwnKeys)
                if (String.Equals(own, path, StringComparison.Ordinal)) return;
            throw new InvalidOperationException("Chave de registro fora do escopo deste instalador.");
        }
    }

    internal sealed class SetupPaths
    {
        internal readonly string Root;
        internal readonly string Dll;
        internal readonly string Setup;
        private readonly string localData;

        internal SetupPaths()
        {
            localData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (String.IsNullOrWhiteSpace(localData) || !Path.IsPathRooted(localData))
                throw new InvalidOperationException("A pasta de dados locais do usuário não está disponível.");
            localData = Path.GetFullPath(localData).TrimEnd(Path.DirectorySeparatorChar);
            Root = Path.GetFullPath(Path.Combine(localData, "MelloCharts"));
            Dll = Path.Combine(Root, SetupConstants.DllName);
            Setup = Path.Combine(Root, SetupConstants.SetupName);
            ValidateRoot();
        }

        internal void ValidateRoot()
        {
            string expected = Path.GetFullPath(Path.Combine(localData, "MelloCharts"));
            if (!Path.IsPathRooted(Root) || !String.Equals(Path.GetFullPath(Root), expected, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Pasta de instalação inválida.");
            if (File.Exists(Root))
                throw new IOException("Existe um arquivo no lugar da pasta de instalação: " + Root);
            if (Directory.Exists(Root) && (File.GetAttributes(Root) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("A pasta de instalação é um link ou ponto de redirecionamento. Operação cancelada.");
        }

        internal void ValidateFile(string path)
        {
            ValidateRoot();
            string full = Path.GetFullPath(path);
            if (!Path.IsPathRooted(path) ||
                (!String.Equals(full, Dll, StringComparison.OrdinalIgnoreCase) &&
                 !String.Equals(full, Setup, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("Arquivo fora do escopo deste instalador.");
            if (Directory.Exists(full)) throw new IOException("O caminho esperado para um arquivo é uma pasta: " + full);
            if (File.Exists(full) && (File.GetAttributes(full) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("O arquivo é um link ou ponto de redirecionamento. Operação cancelada: " + full);
        }

        internal bool RunningFromInstalledSetup()
        {
            return String.Equals(Path.GetFullPath(Application.ExecutablePath), Setup, StringComparison.OrdinalIgnoreCase);
        }

        internal void DeleteKnownFile(string path)
        {
            // Validate the absolute, fixed derived path immediately before every deletion.
            ValidateFile(path);
            if (File.Exists(path)) File.Delete(path);
        }

        internal void DeleteEmptyRoot()
        {
            ValidateRoot();
            if (Directory.Exists(Root) && Directory.GetFileSystemEntries(Root).Length == 0)
                Directory.Delete(Root, false);
        }
    }

    internal sealed class SetupWindow : Form
    {
        private readonly SetupPaths paths;
        private readonly Button install;
        private readonly Button remove;
        private readonly Button close;
        private readonly TextBox output;

        internal SetupWindow(bool uninstallMode)
        {
            paths = new SetupPaths();
            Text = SetupConstants.ProductName + " " + SetupConstants.DisplayVersion;
            ClientSize = new Size(670, 466);
            MinimumSize = new Size(686, 505);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;

            Label heading = new Label();
            heading.Text = "Mello Charts Desktop 0.5.0";
            heading.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
            heading.SetBounds(24, 20, 625, 40);
            Controls.Add(heading);

            Label detail = new Label();
            detail.Text = "Suplemento experimental para PowerPoint no Windows (32 e 64 bits).\r\n" +
                "Instalador sem assinatura digital. Requer .NET Framework 4.8.\r\n\r\n" +
                "Instalação somente para o usuário atual, sem administrador.\r\n" +
                "Feche o PowerPoint antes de instalar, atualizar ou remover.";
            detail.SetBounds(26, 72, 620, 108);
            detail.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(detail);

            Label destination = new Label();
            destination.Text = "Pasta: " + paths.Root;
            destination.AutoEllipsis = true;
            destination.SetBounds(26, 186, 620, 36);
            destination.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(destination);

            install = MakeButton("Instalar", 26, 228);
            remove = MakeButton("Remover", 174, 228);
            Button verify = MakeButton("Verificar", 322, 228);
            verify.Click += delegate { output.Text = MelloCharts.InstallationHealth.Collect(SetupConstants.DisplayVersion); };
            close = MakeButton("Fechar", 522, 228);
            close.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            install.Click += delegate { RunOperation(true); };
            remove.Click += delegate { RunOperation(false); };
            close.Click += delegate { Close(); };
            CancelButton = close;

            output = new TextBox();
            output.Multiline = true;
            output.ReadOnly = true;
            output.ScrollBars = ScrollBars.Vertical;
            output.BackColor = Color.FromArgb(246, 248, 250);
            output.BorderStyle = BorderStyle.FixedSingle;
            output.SetBounds(26, 288, 620, 152);
            output.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            Controls.Add(output);
            output.Text = uninstallMode
                ? "Modo de remoção selecionado. Clique em Remover para continuar.\r\nNenhuma alteração foi feita ao abrir esta janela."
                : "Clique em Instalar para instalar ou atualizar o suplemento.\r\nNenhuma alteração foi feita ao abrir esta janela.";
            Shown += delegate { if (uninstallMode) remove.Focus(); else close.Focus(); };
        }

        private Button MakeButton(string text, int x, int y)
        {
            Button button = new Button();
            button.Text = text;
            button.SetBounds(x, y, 124, 38);
            button.FlatStyle = FlatStyle.System;
            Controls.Add(button);
            return button;
        }

        private void Log(string message)
        {
            output.AppendText(message + "\r\n");
            output.SelectionStart = output.TextLength;
            output.ScrollToCaret();
            output.Refresh();
        }

        private void RunOperation(bool installing)
        {
            install.Enabled = remove.Enabled = close.Enabled = false;
            UseWaitCursor = true;
            output.Clear();
            InstallationTransaction transaction = null;
            try
            {
                SetupConstants.RequirePowerPointClosed();
                paths.ValidateRoot();
                if (installing)
                {
                    byte[] payload = ReadPayload();
                    // Read the setup before mutations; no network, elevation, or child processes.
                    byte[] installerBytes = paths.RunningFromInstalledSetup()
                        ? null : File.ReadAllBytes(Application.ExecutablePath);
                    transaction = new InstallationTransaction(paths);
                    SetupConstants.RequirePowerPointClosed();
                    Directory.CreateDirectory(paths.Root);
                    paths.ValidateRoot();
                    Log("Gravando os arquivos do suplemento...");
                    WriteKnownFile(paths, paths.Dll, payload);

                    // This reads assembly metadata only. It never instantiates or executes the DLL.
                    AssemblyName assembly = AssemblyName.GetAssemblyName(paths.Dll);
                    if (!String.Equals(assembly.Name, "MelloCharts.AddIn", StringComparison.Ordinal) ||
                        assembly.Version == null || assembly.Version.ToString() != SetupConstants.AssemblyVersion)
                        throw new InvalidDataException("O suplemento embutido não corresponde ao MelloCharts.AddIn 0.5.0.0.");
                    if (installerBytes != null) WriteKnownFile(paths, paths.Setup, installerBytes);
                    SetupConstants.RequirePowerPointClosed();
                    Log("Registrando o suplemento somente para este usuário...");
                    Register(paths, assembly.FullName);
                    transaction.Commit();
                    Log("Instalação concluída. Abra o PowerPoint e procure a guia Mello Charts.");
                    Log("Para remover, use este instalador ou Aplicativos nas Configurações do Windows.");
                    MessageBox.Show(this, "Mello Charts Desktop 0.5.0 instalado para este usuário.\r\n" +
                        "Abra o PowerPoint para usar o suplemento.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    transaction = new InstallationTransaction(paths);
                    SetupConstants.RequirePowerPointClosed();
                    Log("Removendo os registros próprios do suplemento...");
                    Unregister();
                    paths.DeleteKnownFile(paths.Dll);
                    bool keepSelf = paths.RunningFromInstalledSetup();
                    if (!keepSelf) paths.DeleteKnownFile(paths.Setup);
                    transaction.Commit();
                    string residue = "";
                    try
                    {
                        paths.DeleteEmptyRoot();
                        if (Directory.Exists(paths.Root))
                            residue = keepSelf
                                ? "\r\nO instalador em execução permanece na pasta. Após fechar esta janela, você pode apagar manualmente " + paths.Setup + "."
                                : "\r\nA pasta contém outros arquivos e foi preservada: " + paths.Root;
                    }
                    catch (Exception ex)
                    {
                        residue = "\r\nA pasta foi preservada: " + paths.Root + ". " + ex.Message;
                    }
                    Log("Suplemento removido para este usuário." + residue);
                    MessageBox.Show(this, "Mello Charts Desktop removido para este usuário." + residue,
                        Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                if (transaction != null && !transaction.IsCommitted)
                {
                    Log("A operação falhou. Restaurando os arquivos e registros anteriores...");
                    List<string> failures = transaction.Rollback();
                    if (failures.Count == 0)
                        message += "\r\n\r\nOs arquivos e registros anteriores foram restaurados.";
                    else
                        message += "\r\n\r\nA restauração não pôde ser concluída:\r\n" + String.Join("\r\n", failures.ToArray()) +
                            "\r\nMantenha o PowerPoint fechado e tente instalar novamente.";
                }
                Log("Operação não concluída. " + message);
                MessageBox.Show(this, message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
                install.Enabled = remove.Enabled = close.Enabled = true;
            }
        }

        private static byte[] ReadPayload()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(SetupConstants.ResourceName))
            {
                if (stream == null) throw new InvalidDataException("O pacote não contém o suplemento embutido.");
                using (MemoryStream buffer = new MemoryStream())
                {
                    stream.CopyTo(buffer);
                    if (buffer.Length == 0) throw new InvalidDataException("O suplemento embutido está vazio.");
                    return buffer.ToArray();
                }
            }
        }

        internal static void WriteKnownFile(SetupPaths paths, string path, byte[] bytes)
        {
            paths.ValidateFile(path);
            if (File.Exists(path) && SameBytes(File.ReadAllBytes(path), bytes)) return;
            File.WriteAllBytes(path, bytes);
        }

        internal static bool SameBytes(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        private static void Register(SetupPaths paths, string assemblyFullName)
        {
            string codeBase = new Uri(paths.Dll, UriKind.Absolute).AbsoluteUri;
            long bytes = new FileInfo(paths.Dll).Length + new FileInfo(paths.Setup).Length;
            foreach (RegistryView view in SetupConstants.RegistryViews())
            {
                using (RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    // Every replaced tree belongs exclusively to this add-in and is backed up first.
                    foreach (string path in SetupConstants.OwnKeys)
                    {
                        SetupConstants.ValidateRegistryPath(path);
                        currentUser.DeleteSubKeyTree(path, false);
                    }
                    using (RegistryKey clsid = currentUser.CreateSubKey(SetupConstants.ClassKey))
                    {
                        clsid.SetValue("", SetupConstants.ProductName, RegistryValueKind.String);
                        using (RegistryKey progId = clsid.CreateSubKey("ProgId"))
                            progId.SetValue("", SetupConstants.ProgId, RegistryValueKind.String);
                        using (RegistryKey server = clsid.CreateSubKey("InprocServer32"))
                        {
                            server.SetValue("", "mscoree.dll", RegistryValueKind.String);
                            server.SetValue("ThreadingModel", "Both", RegistryValueKind.String);
                            SetManagedServerValues(server, assemblyFullName, codeBase);
                            using (RegistryKey version = server.CreateSubKey(SetupConstants.AssemblyVersion))
                                SetManagedServerValues(version, assemblyFullName, codeBase);
                        }
                        using (RegistryKey category = clsid.CreateSubKey("Implemented Categories\\" + SetupConstants.ManagedCategory)) { }
                    }
                    using (RegistryKey progId = currentUser.CreateSubKey(SetupConstants.ProgIdKey))
                    {
                        progId.SetValue("", SetupConstants.ProductName, RegistryValueKind.String);
                        using (RegistryKey clsid = progId.CreateSubKey("CLSID"))
                            clsid.SetValue("", SetupConstants.ClassId, RegistryValueKind.String);
                    }
                    using (RegistryKey office = currentUser.CreateSubKey(SetupConstants.AddInKey))
                    {
                        office.SetValue("FriendlyName", SetupConstants.ProductName, RegistryValueKind.String);
                        office.SetValue("Description", "Gráficos de cascata, build-up e comparação para PowerPoint.", RegistryValueKind.String);
                        office.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                        office.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                    }
                    using (RegistryKey uninstall = currentUser.CreateSubKey(SetupConstants.UninstallKey))
                    {
                        uninstall.SetValue("DisplayName", SetupConstants.ProductName, RegistryValueKind.String);
                        uninstall.SetValue("DisplayVersion", SetupConstants.DisplayVersion, RegistryValueKind.String);
                        uninstall.SetValue("Publisher", "Mello Charts", RegistryValueKind.String);
                        uninstall.SetValue("InstallLocation", paths.Root, RegistryValueKind.String);
                        uninstall.SetValue("DisplayIcon", "\"" + paths.Setup + "\",0", RegistryValueKind.String);
                        uninstall.SetValue("UninstallString", "\"" + paths.Setup + "\" --uninstall", RegistryValueKind.String);
                        uninstall.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        uninstall.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                        uninstall.SetValue("EstimatedSize", (int)Math.Min(Int32.MaxValue, (bytes + 1023L) / 1024L), RegistryValueKind.DWord);
                    }
                }
            }
        }

        private static void SetManagedServerValues(RegistryKey key, string assemblyFullName, string codeBase)
        {
            key.SetValue("Class", SetupConstants.ProgId, RegistryValueKind.String);
            key.SetValue("Assembly", assemblyFullName, RegistryValueKind.String);
            key.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
            key.SetValue("CodeBase", codeBase, RegistryValueKind.String);
        }

        private static void Unregister()
        {
            foreach (RegistryView view in SetupConstants.RegistryViews())
            {
                using (RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                    foreach (string path in SetupConstants.OwnKeys)
                    {
                        SetupConstants.ValidateRegistryPath(path);
                        currentUser.DeleteSubKeyTree(path, false);
                    }
            }
        }
    }

    // In-memory snapshots keep previous versions and values available until the operation commits.
    // This handles ordinary write/registration failures; it is not a power-loss recovery journal.
    internal sealed class InstallationTransaction
    {
        private readonly SetupPaths paths;
        private readonly bool directoryExisted;
        private readonly List<FileSnapshot> files = new List<FileSnapshot>();
        private readonly List<RegistrySnapshot> keys = new List<RegistrySnapshot>();
        internal bool IsCommitted { get; private set; }

        internal InstallationTransaction(SetupPaths target)
        {
            paths = target;
            paths.ValidateRoot();
            directoryExisted = Directory.Exists(paths.Root);
            files.Add(new FileSnapshot(paths, paths.Dll));
            files.Add(new FileSnapshot(paths, paths.Setup));
            foreach (RegistryView view in SetupConstants.RegistryViews())
                foreach (string key in SetupConstants.OwnKeys)
                    keys.Add(new RegistrySnapshot(view, key));
        }

        internal void Commit() { IsCommitted = true; }

        internal List<string> Rollback()
        {
            List<string> errors = new List<string>();
            foreach (FileSnapshot file in files)
            {
                try { file.Restore(paths); }
                catch (Exception ex) { errors.Add(file.Path + ": " + ex.Message); }
            }
            // 32/64-bit HKCU branches can share underlying keys on some Windows releases.
            // Restoring in reverse order also restores their initial shared state.
            for (int i = keys.Count - 1; i >= 0; i--)
            {
                try { keys[i].Restore(); }
                catch (Exception ex) { errors.Add(keys[i].Description + ": " + ex.Message); }
            }
            if (!directoryExisted)
            {
                try { paths.DeleteEmptyRoot(); }
                catch (Exception ex) { errors.Add(paths.Root + ": " + ex.Message); }
            }
            return errors;
        }
    }

    internal sealed class FileSnapshot
    {
        internal readonly string Path;
        private readonly byte[] data;
        private readonly FileAttributes attributes;
        private readonly DateTime lastWrite;

        internal FileSnapshot(SetupPaths paths, string path)
        {
            paths.ValidateFile(path);
            Path = path;
            if (File.Exists(path))
            {
                data = File.ReadAllBytes(path);
                attributes = File.GetAttributes(path);
                lastWrite = File.GetLastWriteTimeUtc(path);
            }
        }

        internal void Restore(SetupPaths paths)
        {
            paths.ValidateFile(Path);
            if (data == null) { paths.DeleteKnownFile(Path); return; }
            if (File.Exists(Path) && SetupWindow.SameBytes(File.ReadAllBytes(Path), data)) return;
            Directory.CreateDirectory(paths.Root);
            SetupWindow.WriteKnownFile(paths, Path, data);
            File.SetLastWriteTimeUtc(Path, lastWrite);
            File.SetAttributes(Path, attributes);
        }
    }

    internal sealed class RegistrySnapshot
    {
        private readonly RegistryView view;
        private readonly string path;
        private readonly RegistryNode tree;
        internal string Description { get { return "HKCU (" + view + ")\\" + path; } }

        internal RegistrySnapshot(RegistryView targetView, string targetPath)
        {
            SetupConstants.ValidateRegistryPath(targetPath);
            view = targetView;
            path = targetPath;
            using (RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            using (RegistryKey key = currentUser.OpenSubKey(path, false))
                tree = key == null ? null : RegistryNode.Read(key);
        }

        internal void Restore()
        {
            SetupConstants.ValidateRegistryPath(path);
            using (RegistryKey currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                currentUser.DeleteSubKeyTree(path, false);
                if (tree != null)
                    using (RegistryKey key = currentUser.CreateSubKey(path)) tree.Write(key);
            }
        }
    }

    internal sealed class RegistryNode
    {
        private readonly Dictionary<string, RegistryValue> values = new Dictionary<string, RegistryValue>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, RegistryNode> children = new Dictionary<string, RegistryNode>(StringComparer.OrdinalIgnoreCase);

        internal static RegistryNode Read(RegistryKey key)
        {
            RegistryNode result = new RegistryNode();
            foreach (string name in key.GetValueNames())
            {
                object value = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                if (value == null) throw new IOException("Não foi possível salvar um valor do registro antes da operação.");
                result.values.Add(name, new RegistryValue(value, key.GetValueKind(name)));
            }
            foreach (string name in key.GetSubKeyNames())
                using (RegistryKey child = key.OpenSubKey(name, false))
                {
                    if (child == null) throw new IOException("A chave do registro mudou durante a leitura. Tente novamente.");
                    result.children.Add(name, Read(child));
                }
            return result;
        }

        internal void Write(RegistryKey key)
        {
            foreach (KeyValuePair<string, RegistryValue> value in values)
                key.SetValue(value.Key, value.Value.Data, value.Value.Kind);
            foreach (KeyValuePair<string, RegistryNode> child in children)
                using (RegistryKey subkey = key.CreateSubKey(child.Key)) child.Value.Write(subkey);
        }
    }

    internal sealed class RegistryValue
    {
        internal readonly object Data;
        internal readonly RegistryValueKind Kind;
        internal RegistryValue(object data, RegistryValueKind kind) { Data = data; Kind = kind; }
    }
}
