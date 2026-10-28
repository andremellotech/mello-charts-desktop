using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace MelloCharts {
    public sealed class ChartColorForm:Form {
        private readonly CheckBox enabled=new CheckBox();
        private readonly Dictionary<string,TextBox> inputs=new Dictionary<string,TextBox>();
        private readonly Dictionary<string,Button> swatches=new Dictionary<string,Button>();
        private readonly Label error=new Label();
        private readonly TabControl tabs=new TabControl();
        private readonly ChartColors defaults;
        public ChartColors Result {get;private set;}
        public ChartColorForm(ChartDocument document) {
            defaults=ChartCore.ColorDefaults(document);ChartColors effective=ChartCore.EffectiveColors(document);
            Text="Mello Charts · Cores personalizadas";Font=new Font("Segoe UI",9);BackColor=Color.FromArgb(239,243,246);
            AutoScaleMode=AutoScaleMode.Dpi;StartPosition=FormStartPosition.CenterParent;ShowInTaskbar=false;
            Rectangle area=Screen.PrimaryScreen.WorkingArea;Size=new Size(Math.Min(780,area.Width-32),Math.Min(760,area.Height-32));MinimumSize=new Size(600,460);
            TableLayoutPanel root=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(18),ColumnCount=1,RowCount=4};
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,88));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,60));root.RowStyles.Add(new RowStyle(SizeType.Absolute,42));Controls.Add(root);
            Panel heading=new Panel {Dock=DockStyle.Fill};
            enabled.Text="Usar cores personalizadas";enabled.Font=new Font(Font,FontStyle.Bold);enabled.Dock=DockStyle.Top;enabled.Height=30;enabled.Checked=document.customColors!=null;
            heading.Controls.Add(new Label {Text="Qualquer cor RGB: digite #RRGGBB, #RGB ou rgb(18, 52, 86).\nEm Escolher, ajuste os valores de vermelho, verde e azul. O texto se adapta ao contraste.",Dock=DockStyle.Bottom,Height=50,ForeColor=Color.FromArgb(65,85,102)});
            heading.Controls.Add(enabled);root.Controls.Add(heading,0,0);tabs.Dock=DockStyle.Fill;root.Controls.Add(tabs,0,1);
            FlowLayoutPanel general=Page("Séries e fundo"),categories=Page("Pizza e rosca");
            string[] keys={"actual","forecast","baseline","increase","decrease","background"};
            string[] captions={"Realizado / base / total","Projetado / linha","Baseline / acumulado","Aumento na cascata","Redução na cascata","Fundo do gráfico"};
            for(int i=0;i<keys.Length;i++)Row(general,keys[i],captions[i],(string)typeof(ChartColors).GetField(keys[i]).GetValue(effective));
            for(int i=0;i<12;i++)Row(categories,"category"+i,(i+1)+". "+(i<document.rows.Count?document.rows[i].label:"Categoria"),effective.categories[i]);
            error.Dock=DockStyle.Fill;error.ForeColor=Color.FromArgb(154,55,36);error.Padding=new Padding(0,8,0,0);root.Controls.Add(error,0,2);
            FlowLayoutPanel actions=new FlowLayoutPanel {Dock=DockStyle.Fill,FlowDirection=FlowDirection.RightToLeft,WrapContents=false};
            Button accept=new Button {Text="Aplicar cores",Width=125,Height=32,BackColor=Color.FromArgb(36,79,115),ForeColor=Color.White,FlatStyle=FlatStyle.Flat};
            accept.Click+=(s,e)=>TryAccept();actions.Controls.Add(accept);
            Button cancel=new Button {Text="Cancelar",Width=90,Height=32,DialogResult=DialogResult.Cancel};actions.Controls.Add(cancel);
            Button reset=new Button {Text="Restaurar padrão",Width=135,Height=32};reset.Click+=(s,e)=>ResetDefaults();actions.Controls.Add(reset);root.Controls.Add(actions,0,3);
            AcceptButton=accept;CancelButton=cancel;enabled.CheckedChanged+=(s,e)=>{tabs.Enabled=enabled.Checked;error.Text="";};tabs.Enabled=enabled.Checked;
        }
        private FlowLayoutPanel Page(string title) {
            TabPage page=new TabPage(title){BackColor=Color.White,Padding=new Padding(10)};tabs.TabPages.Add(page);
            FlowLayoutPanel panel=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoScroll=true,FlowDirection=FlowDirection.TopDown,WrapContents=false};page.Controls.Add(panel);return panel;
        }
        private void Row(FlowLayoutPanel parent,string key,string caption,string color) {
            TableLayoutPanel row=new TableLayoutPanel {Width=660,Height=46,ColumnCount=3,Margin=new Padding(0,0,0,4)};
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,48));row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,32));row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,20));
            row.Controls.Add(new Label {Text=caption,Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft,AutoEllipsis=true},0,0);
            TextBox input=new TextBox {Text=color,Dock=DockStyle.Fill,MaxLength=60,Margin=new Padding(4,10,8,0),AccessibleName=caption+" em HEX ou RGB"};inputs.Add(key,input);row.Controls.Add(input,1,0);
            Button picker=new Button {Text="Escolher…",Dock=DockStyle.Fill,Margin=new Padding(0,6,6,6),FlatStyle=FlatStyle.Flat,AccessibleName="Escolher cor de "+caption};swatches.Add(key,picker);row.Controls.Add(picker,2,0);
            input.TextChanged+=(s,e)=>UpdateSwatch(key);picker.Click+=(s,e)=>Pick(key);UpdateSwatch(key);parent.Controls.Add(row);
            parent.SizeChanged+=(s,e)=>row.Width=Math.Max(460,parent.ClientSize.Width-SystemInformation.VerticalScrollBarWidth-6);
        }
        private void UpdateSwatch(string key) {
            try {string color=ColorTools.Normalize(inputs[key].Text);swatches[key].BackColor=ColorTranslator.FromHtml(color);swatches[key].ForeColor=ColorTranslator.FromHtml(ColorTools.Ink(color));}
            catch(ArgumentException){swatches[key].BackColor=Color.White;swatches[key].ForeColor=Color.FromArgb(154,55,36);}
        }
        private void Pick(string key) {
            string initial;try{initial=ColorTools.Normalize(inputs[key].Text);}catch(ArgumentException){initial="#FFFFFF";}
            using(ColorDialog dialog=new ColorDialog {FullOpen=true,AnyColor=true,Color=ColorTranslator.FromHtml(initial)}) {
                if(dialog.ShowDialog(this)==DialogResult.OK)inputs[key].Text="#"+dialog.Color.R.ToString("X2")+dialog.Color.G.ToString("X2")+dialog.Color.B.ToString("X2");
            }
        }
        private void ResetDefaults() {
            foreach(var pair in inputs)pair.Value.Text=pair.Key.StartsWith("category")?defaults.categories[Int32.Parse(pair.Key.Substring(8))]:(string)typeof(ChartColors).GetField(pair.Key).GetValue(defaults);
            enabled.Checked=false;error.Text="";
        }
        private bool TryAccept() {
            try {
                ChartColors result=null;
                if(enabled.Checked) {
                    result=new ChartColors();bool changed=false;
                    foreach(string key in new[]{"actual","forecast","baseline","increase","decrease","background"}) {
                        string value=ColorTools.Normalize(inputs[key].Text),original=(string)typeof(ChartColors).GetField(key).GetValue(defaults);
                        if(value!=original){typeof(ChartColors).GetField(key).SetValue(result,value);changed=true;}
                    }
                    List<string> categories=Enumerable.Range(0,12).Select(i=>ColorTools.Normalize(inputs["category"+i].Text)).ToList();
                    if(!categories.SequenceEqual(defaults.categories)){result.categories=categories;changed=true;}
                    if(!changed)result=null;
                }
                Result=result;DialogResult=DialogResult.OK;Close();return true;
            } catch(ArgumentException ex){error.Text=ex.Message;return false;}
        }
    }
}
