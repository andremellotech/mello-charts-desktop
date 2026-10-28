using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MelloCharts
{
    public sealed class WorkbookImportForm:Form
    {
        readonly string chartType;
        readonly TextBox filePath=new TextBox(),range=new TextBox();
        readonly ComboBox sheet=new ComboBox();
        readonly DataGridView table=new DataGridView();
        readonly ChartPreview chart=new ChartPreview();
        readonly Button confirm=new Button(),preview=new Button();
        readonly Label status=new Label();
        ChartDocument candidate;
        public ChartDocument Result {get;private set;}
        public WorkbookImportForm(string type)
        {
            ChartDefinition definition=ChartCatalog.Get(type);chartType=type;
            Text="Importar Excel · Mello Charts";StartPosition=FormStartPosition.CenterParent;
            AutoScaleMode=AutoScaleMode.Dpi;ClientSize=new Size(1040,780);MinimumSize=new Size(880,680);
            Font=new Font("Segoe UI",10);BackColor=Color.FromArgb(244,247,249);ShowIcon=false;ShowInTaskbar=false;
            TableLayoutPanel root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=6,Padding=new Padding(22,16,22,16)};
            foreach(int height in new[]{62,54,72,92})root.RowStyles.Add(new RowStyle(SizeType.Absolute,height));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,78));Controls.Add(root);
            Label heading=new Label{Text="Conectar dados do Excel",Font=new Font(Font.FontFamily,19,FontStyle.Bold),ForeColor=Color.FromArgb(32,55,77),Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft};root.Controls.Add(heading,0,0);
            TableLayoutPanel fileRow=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=1,Margin=new Padding(0)};fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));fileRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,158));
            filePath.ReadOnly=true;filePath.Dock=DockStyle.Fill;filePath.AccessibleName="Arquivo XLSX local";filePath.Margin=new Padding(0,6,10,8);fileRow.Controls.Add(filePath,0,0);
            Button browse=Button("Escolher XLSX…",false);browse.Dock=DockStyle.Fill;browse.Margin=new Padding(0,2,0,12);browse.Click+=(s,e)=>ChooseFile();fileRow.Controls.Add(browse,1,0);root.Controls.Add(fileRow,0,1);
            TableLayoutPanel selectors=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,RowCount=2,Margin=new Padding(0)};selectors.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));selectors.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));selectors.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,158));selectors.RowStyles.Add(new RowStyle(SizeType.Absolute,25));selectors.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            selectors.Controls.Add(Label("Planilha"),0,0);selectors.Controls.Add(Label("Intervalo A1, incluindo cabeçalho"),1,0);
            sheet.DropDownStyle=ComboBoxStyle.DropDownList;sheet.Dock=DockStyle.Fill;sheet.AccessibleName="Planilha do arquivo";sheet.Margin=new Padding(0,0,10,10);selectors.Controls.Add(sheet,0,1);
            range.Dock=DockStyle.Fill;range.AccessibleName="Intervalo A1 incluindo cabeçalho";range.Margin=new Padding(0,0,10,10);selectors.Controls.Add(range,1,1);
            preview.Text="Ver prévia";StyleButton(preview,false);preview.Dock=DockStyle.Fill;preview.Margin=new Padding(0,0,0,10);preview.Enabled=false;preview.Click+=(s,e)=>LoadPreview();selectors.Controls.Add(preview,2,1);root.Controls.Add(selectors,0,2);
            Label help=Label("Inclua o cabeçalho na primeira linha da seleção. Ex.: H500:K510. Deixe vazio para detectar a área usada.\nCabeçalho para "+definition.Name+": "+definition.Header+"\nCategorias: datas padrão, mmm/yyyy, dd/MM/yyyy e códigos 000000. Outros formatos: converta em texto no Excel.\nDados salvos no arquivo. Para trazer alterações futuras, salve no Excel e use Atualizar do Excel explicitamente.");help.Dock=DockStyle.Fill;help.ForeColor=Color.FromArgb(72,96,112);root.Controls.Add(help,0,3);
            TabControl tabs=new TabControl{Dock=DockStyle.Fill,Margin=new Padding(0)};TabPage chartTab=new TabPage("Prévia do gráfico"),tableTab=new TabPage("Tabela selecionada");chart.Dock=DockStyle.Fill;chartTab.Controls.Add(chart);table.Dock=DockStyle.Fill;table.ReadOnly=true;table.AllowUserToAddRows=false;table.AllowUserToDeleteRows=false;table.AllowUserToResizeRows=false;table.RowHeadersVisible=false;table.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;table.BackgroundColor=Color.White;table.BorderStyle=BorderStyle.None;table.ColumnHeadersHeight=34;table.RowTemplate.Height=28;table.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(244,248,250);tableTab.Controls.Add(table);tabs.TabPages.Add(chartTab);tabs.TabPages.Add(tableTab);root.Controls.Add(tabs,0,4);
            TableLayoutPanel footer=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,RowCount=1,Margin=new Padding(0,12,0,0)};footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,177));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,100));
            status.Dock=DockStyle.Fill;status.TextAlign=ContentAlignment.MiddleLeft;status.ForeColor=Color.FromArgb(67,97,111);status.Text="Escolha um XLSX local de até 10 MB para selecionar a aba e o intervalo.";status.AccessibleName="Estado da importação";footer.Controls.Add(status,0,0);
            confirm.Text="Importar e vincular";StyleButton(confirm,true);confirm.Dock=DockStyle.Fill;confirm.Margin=new Padding(10,8,6,8);confirm.Enabled=false;confirm.Click+=(s,e)=>AcceptImport();footer.Controls.Add(confirm,1,0);
            Button cancel=Button("Cancelar",false);cancel.Dock=DockStyle.Fill;cancel.Margin=new Padding(0,8,0,8);cancel.DialogResult=DialogResult.Cancel;footer.Controls.Add(cancel,2,0);CancelButton=cancel;root.Controls.Add(footer,0,5);
            sheet.SelectedIndexChanged+=(s,e)=>InvalidatePreview();range.TextChanged+=(s,e)=>InvalidatePreview();
        }
        static Label Label(string text){return new Label{Text=text,AutoSize=false,Dock=DockStyle.Fill,ForeColor=Color.FromArgb(40,70,85)};}
        static Button Button(string text,bool primary){Button button=new Button{Text=text};StyleButton(button,primary);return button;}
        static void StyleButton(Button button,bool primary){button.FlatStyle=FlatStyle.Flat;button.FlatAppearance.BorderColor=Color.FromArgb(208,221,226);button.BackColor=primary?Color.FromArgb(36,79,115):Color.White;button.ForeColor=primary?Color.White:Color.FromArgb(32,55,77);button.UseVisualStyleBackColor=false;}
        void InvalidatePreview(){candidate=null;confirm.Enabled=false;chart.Scene=null;chart.Invalidate();table.Rows.Clear();table.Columns.Clear();status.ForeColor=Color.FromArgb(67,97,111);status.Text="Inclua o cabeçalho no intervalo e clique em Ver prévia para validar os dados salvos.";}
        void ChooseFile()
        {
            using(OpenFileDialog dialog=new OpenFileDialog{Filter="Pasta de trabalho Excel (*.xlsx)|*.xlsx",Title="Escolher arquivo XLSX local salvo",CheckFileExists=true,Multiselect=false}){
                if(dialog.ShowDialog(this)!=DialogResult.OK)return;SetPath(dialog.FileName);
            }
        }
        public void SetPath(string path)
        {
            InvalidatePreview();filePath.Text="";sheet.Items.Clear();preview.Enabled=false;
            try{IList<string> sheets=WorkbookData.ListSheets(path);filePath.Text=Path.GetFullPath(path);foreach(string name in sheets)sheet.Items.Add(name);if(sheet.Items.Count>0)sheet.SelectedIndex=0;range.Clear();preview.Enabled=true;}
            catch(Exception e){ShowError(e);}
        }
        void LoadPreview()
        {
            InvalidatePreview();try{
                WorkbookSource link;List<List<string>> values=WorkbookData.ReadTable(filePath.Text,Convert.ToString(sheet.SelectedItem),range.Text,chartType,out link);
                ChartDocument imported=WorkbookData.FromTable(values,chartType,link);Scene scene=ChartCore.BuildScene(imported);
                for(int col=0;col<values[0].Count;col++)table.Columns.Add("c"+col,values[0][col]);for(int row=1;row<values.Count;row++)table.Rows.Add(values[row].ToArray());
                candidate=imported;chart.Scene=scene;chart.Invalidate();confirm.Enabled=true;status.ForeColor=Color.FromArgb(39,107,90);status.Text=link.sheet+" · "+link.range+" · "+imported.rows.Count+" linhas. Prévia validada dos dados salvos.";
            }catch(Exception e){ShowError(e);}
        }
        void AcceptImport(){if(candidate==null)return;Result=ChartCore.Clone(candidate);DialogResult=DialogResult.OK;Close();}
        void ShowError(Exception e){candidate=null;confirm.Enabled=false;status.ForeColor=Color.FromArgb(157,57,53);status.Text=e.Message;}
    }
}
