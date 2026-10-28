using System;
using System.Drawing;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace MelloCharts {
  public sealed class EditorForm:Form {
    private object application;private ChartContext context;
    private readonly List<string> undoStates=new List<string>(),redoStates=new List<string>();private string historyCurrent;private bool historyRestoring,invalidEdit;
    private readonly Button undoButton=new Button(),redoButton=new Button(),refreshDataButton=new Button();private ChartReference target;private ChartDocument document,lastValidDocument;private string savedState;private bool closingFromHost,acceptedClose,snapshotDirty,cleanupQueued;private string recoveryFolder=null;private System.Threading.SynchronizationContext ownerUi;
    public event EventHandler OwnerWindowClosed;
    private readonly TextBox title=new TextBox(),subtitle=new TextBox(),source=new TextBox(),unit=new TextBox(),seriesActual=new TextBox(),seriesForecast=new TextBox(),seriesBaseline=new TextBox();
    private readonly ComboBox palette=new ComboBox(),decimals=new ComboBox(),style=new ComboBox(),fontFamily=new ComboBox();
    private readonly CheckBox labels=new CheckBox(),baseline=new CheckBox(),total=new CheckBox(),compact=new CheckBox(),gridLines=new CheckBox();
    private readonly ToolTip tips=new ToolTip();
    private ChartColors customColors;
    private readonly Button colorButton=new Button();
    private readonly DataGridView grid=new DataGridView();private readonly ChartPreview preview=new ChartPreview();
    private readonly Label status=new Label(),method=new Label(),typeLabel=new Label();private readonly Button apply=new Button();private bool loading=false,valid=false;
    private static readonly CultureInfo Pt=CultureInfo.GetCultureInfo("pt-BR");
    public EditorForm(object app,ChartDocument doc,ChartReference reference){
      application=app;target=reference;document=ChartCore.Clone(doc);
      context=app==null?null:PowerPointBridge.CaptureContext(app);savedState=ChartCore.Serialize(document);lastValidDocument=ChartCore.Clone(document);historyCurrent=savedState;
      if(context!=null && reference!=null && context.SlideId!=reference.SlideId)throw new InvalidOperationException("Selecione novamente o slide do gráfico.");
      Text="Mello Charts · "+(reference==null?"Novo gráfico":"Editar gráfico");Font=new Font("Segoe UI",9);BackColor=Color.FromArgb(239,243,246);StartPosition=FormStartPosition.CenterScreen;
      if(context!=null){Text+=" · slide "+context.SlideNumber;ShowInTaskbar=false;}
      System.Drawing.Rectangle area=Screen.PrimaryScreen.WorkingArea;Size=new Size(Math.Min(1240,area.Width-32),Math.Min(930,area.Height-32));MinimumSize=new Size(900,640);AutoScaleMode=AutoScaleMode.Dpi;
      TableLayoutPanel root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(16)};
      root.RowStyles.Add(new RowStyle(SizeType.Absolute,90));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,54));Controls.Add(root);
      FlowLayoutPanel tools=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=true,AutoScroll=true};
      typeLabel.AutoSize=true;typeLabel.Font=new Font("Segoe UI",11,FontStyle.Bold);typeLabel.ForeColor=Color.FromArgb(28,62,80);typeLabel.Margin=new Padding(0,8,18,0);tools.Controls.Add(typeLabel);
      Tool(tools,"Importar tabela",Import);Tool(tools,"Vincular Excel",LinkExcel);
      ConfigureTool(refreshDataButton,"Atualizar dados",RefreshExcel);tools.Controls.Add(refreshDataButton);
      ConfigureTool(undoButton,"↶ Desfazer",UndoEdit);tools.Controls.Add(undoButton);ConfigureTool(redoButton,"↷ Refazer",RedoEdit);tools.Controls.Add(redoButton);Tool(tools,"Colar do Excel",PasteTable);Tool(tools,"Abrir projeto",OpenProject);Tool(tools,"Salvar projeto",SaveProject);Tool(tools,"Exportar PNG",ExportPng);Tool(tools,"Exportar CSV",ExportCsv);Tool(tools,"Recuperar rascunho",OpenRecovery);root.Controls.Add(tools,0,0);
      TableLayoutPanel body=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=2,Margin=new Padding(0)};
      body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,258));body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));body.RowStyles.Add(new RowStyle(SizeType.Percent,65));body.RowStyles.Add(new RowStyle(SizeType.Percent,35));root.Controls.Add(body,0,1);
      Panel side=new Panel{Dock=DockStyle.Fill,AutoScroll=true,BackColor=Color.White,Padding=new Padding(14,12,14,12),Margin=new Padding(0,0,14,0)};
      FlowLayoutPanel settings=new FlowLayoutPanel{Dock=DockStyle.Top,AutoSize=true,FlowDirection=FlowDirection.TopDown,WrapContents=false,Margin=new Padding(0)};side.Controls.Add(settings);body.Controls.Add(side,0,0);body.SetRowSpan(side,2);
      SideHeading(settings,"APARÊNCIA");
      if(context!=null)settings.Controls.Add(new Label{Text="Vinculado ao slide "+context.SlideNumber+". Você pode navegar no PowerPoint sem alterar esse destino.",Width=202,Height=58,ForeColor=Color.FromArgb(92,111,125),Margin=new Padding(0,0,0,12)});
      style.DropDownStyle=ComboBoxStyle.DropDownList;style.Items.AddRange(new object[]{"Executivo","Essencial","Escuro"});SideField(settings,"Estilo pronto",style,48);
      palette.DropDownStyle=ComboBoxStyle.DropDownList;palette.Items.AddRange(new object[]{"Atlântico · azul e verde","Floresta · verde e sálvia","Grafite · cinza e ardósia"});SideField(settings,"Paleta pronta",palette,48);
      colorButton.Text="Cores personalizadas…";colorButton.Width=202;colorButton.Height=34;colorButton.FlatStyle=FlatStyle.Flat;colorButton.BackColor=Color.White;colorButton.Margin=new Padding(0,0,0,10);colorButton.Click+=(s,e)=>CustomizeColors();settings.Controls.Add(colorButton);
      tips.SetToolTip(colorButton,"Use qualquer cor RGB: seletor, HEX ou valores RGB. As cores são salvas no projeto e no grupo do slide. Escolher outra paleta pronta restaura suas cores.");
      decimals.DropDownStyle=ComboBoxStyle.DropDownList;decimals.Items.AddRange(new object[]{"0","1","2"});SideField(settings,"Casas decimais",decimals,48);
      labels.Text="Mostrar valores";baseline.Text="Mostrar baseline";total.Text="Adicionar total";compact.Text="Abreviar números";gridLines.Text="Linhas de referência";
      foreach(CheckBox check in new[]{labels,baseline,total,compact,gridLines}){check.AutoSize=true;check.Margin=new Padding(0,0,0,4);settings.Controls.Add(check);check.CheckedChanged+=(s,e)=>RefreshChart();}
      tips.SetToolTip(compact,"Abrevia na unidade informada: 1.500 vira 1,5 mil. Os valores da tabela, do CSV e do projeto não mudam.");
      tips.SetToolTip(style,"Executivo destaca o consolidado. Essencial reduz os elementos decorativos. Escuro usa cores com contraste próprio.");
      tips.SetToolTip(gridLines,"Oculta as linhas auxiliares e mantém a referência do zero.");
      fontFamily.DropDownStyle=ComboBoxStyle.DropDownList;fontFamily.Items.AddRange(new object[]{"Arial","Calibri","Aptos","Segoe UI"});SideField(settings,"Fonte tipográfica",fontFamily,48);
      SideHeading(settings,"CONTEÚDO");title.Multiline=true;subtitle.Multiline=true;SideField(settings,"Título",title,68);SideField(settings,"Subtítulo",subtitle,68);SideField(settings,"Unidade dos dados",unit,48);SideField(settings,"Fonte dos dados",source,48);
      SideField(settings,"Primeira série",seriesActual,48);SideField(settings,"Segunda série",seriesForecast,48);SideField(settings,"Série de referência",seriesBaseline,48);foreach(TextBox box in new[]{seriesActual,seriesForecast,seriesBaseline})box.MaxLength=24;
      preview.Dock=DockStyle.Fill;preview.Margin=new Padding(0,0,0,10);body.Controls.Add(preview,1,0);
      TableLayoutPanel dataPanel=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Margin=new Padding(0)};dataPanel.RowStyles.Add(new RowStyle(SizeType.Percent,100));dataPanel.RowStyles.Add(new RowStyle(SizeType.Absolute,56));body.Controls.Add(dataPanel,1,1);
      grid.Dock=DockStyle.Fill;grid.Margin=new Padding(0);grid.BackgroundColor=Color.White;grid.BorderStyle=BorderStyle.None;grid.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.Fill;grid.RowHeadersWidth=36;grid.AllowUserToAddRows=false;grid.AllowUserToDeleteRows=true;grid.SelectionMode=DataGridViewSelectionMode.CellSelect;grid.MultiSelect=true;grid.RowTemplate.Height=29;grid.ColumnHeadersHeight=34;grid.EnableHeadersVisualStyles=false;
      grid.ColumnHeadersDefaultCellStyle.BackColor=Color.FromArgb(222,233,239);grid.ColumnHeadersDefaultCellStyle.ForeColor=Color.FromArgb(40,70,85);grid.DefaultCellStyle.SelectionBackColor=Color.FromArgb(211,234,231);grid.DefaultCellStyle.SelectionForeColor=Color.FromArgb(28,55,70);grid.GridColor=Color.FromArgb(229,236,240);grid.AlternatingRowsDefaultCellStyle.BackColor=Color.FromArgb(248,250,252);dataPanel.Controls.Add(grid,0,0);
      FlowLayoutPanel rowTools=new FlowLayoutPanel{Dock=DockStyle.Fill,WrapContents=false,AutoScroll=true};Tool(rowTools,"+ Categoria",AddRow);Tool(rowTools,"Excluir linha",DeleteRow);method.AutoSize=true;method.MaximumSize=new Size(510,32);method.ForeColor=Color.FromArgb(92,113,126);method.Margin=new Padding(8,9,0,0);rowTools.Controls.Add(method);dataPanel.Controls.Add(rowTools,0,1);
      TableLayoutPanel footer=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=3,Margin=new Padding(0,8,0,0)};footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,177));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,86));
      status.Dock=DockStyle.Fill;status.ForeColor=Color.FromArgb(67,97,111);status.TextAlign=ContentAlignment.MiddleLeft;footer.Controls.Add(status,0,0);
      apply.Text=app==null?"Enviar ao PowerPoint":target==null?"Inserir no slide":"Atualizar gráfico";apply.Dock=DockStyle.Fill;apply.BackColor=Color.FromArgb(36,79,115);apply.ForeColor=Color.White;apply.FlatStyle=FlatStyle.Flat;apply.FlatAppearance.BorderSize=0;apply.Click+=(s,e)=>ApplyChart();footer.Controls.Add(apply,1,0);
      Button close=new Button{Text="Fechar",Dock=DockStyle.Fill,FlatStyle=FlatStyle.Flat,BackColor=Color.White};close.FlatAppearance.BorderColor=Color.FromArgb(208,221,226);close.Click+=(s,e)=>Close();footer.Controls.Add(close,2,0);root.Controls.Add(footer,0,2);
      foreach(TextBox box in new[]{title,subtitle,source,unit,seriesActual,seriesForecast,seriesBaseline})box.TextChanged+=(s,e)=>RefreshChart();foreach(ComboBox combo in new[]{decimals,style,fontFamily})combo.SelectedIndexChanged+=(s,e)=>RefreshChart();
      palette.SelectedIndexChanged+=(s,e)=>{if(!loading)customColors=null;UpdateColorButton();RefreshChart();};
      grid.CellValueChanged+=(s,e)=>RefreshChart();grid.RowsRemoved+=(s,e)=>RefreshChart();grid.CurrentCellDirtyStateChanged+=(s,e)=>{if(grid.IsCurrentCellDirty)grid.CommitEdit(DataGridViewDataErrorContexts.Commit);};grid.DataError+=(s,e)=>{e.ThrowException=false;SetStatus("Confira o tipo e os valores informados na tabela.",true);};
      grid.UserDeletingRow+=(s,e)=>{if(grid.Rows.Count<=1||(document.type=="waterfall"&&e.Row.Index==0)){e.Cancel=true;SetStatus("Mantenha pelo menos uma categoria e a base inicial da cascata.",true);}};
      LoadDocument();
    }
    protected override void Dispose(bool disposing){if(disposing)tips.Dispose();base.Dispose(disposing);}
    public bool HasPendingChanges{get{try{return ChartCore.Serialize(ReadEditor())!=savedState;}catch{return true;}}}
    private void RecoverDraft(){if(snapshotDirty && lastValidDocument!=null)RecoveryStore.SaveTo(recoveryFolder??RecoveryStore.Folder,lastValidDocument);}
    protected override void OnShown(EventArgs e){ownerUi=System.Threading.SynchronizationContext.Current as WindowsFormsSynchronizationContext??new WindowsFormsSynchronizationContext();base.OnShown(e);}
    protected override void WndProc(ref Message message){
      // Destroying an external native owner skips FormClosing/FormClosed. Save only
      // cached valid state here; control HWNDs are being destroyed, so do not read them.
      if(message.Msg==0x0002 && !RecreatingHandle && !Disposing && !closingFromHost && !acceptedClose && !cleanupQueued){
        cleanupQueued=true;try{RecoverDraft();}catch(Exception ex){System.Diagnostics.Trace.TraceError("Mello Charts recovery: "+ex.GetType().Name);}
        if(ownerUi!=null)try{ownerUi.Post(delegate(object ignored){try{EventHandler handler=OwnerWindowClosed;if(handler!=null)handler(this,EventArgs.Empty);}finally{Dispose();}},null);}catch(Exception ex){System.Diagnostics.Trace.TraceError("Mello Charts owner cleanup: "+ex.GetType().Name);}
      }
      base.WndProc(ref message);
    }
    public void CloseFromHost(){try{grid.EndEdit();RefreshChart();RecoverDraft();}catch(Exception ex){System.Diagnostics.Trace.TraceError("Mello Charts recovery: "+ex.GetType().Name);}closingFromHost=true;Close();}
    protected override void OnFormClosing(FormClosingEventArgs e){
      grid.EndEdit();RefreshChart();
      if(!closingFromHost && e.CloseReason!=CloseReason.UserClosing){try{RecoverDraft();}catch(Exception ex){System.Diagnostics.Trace.TraceError("Mello Charts recovery: "+ex.GetType().Name);}closingFromHost=true;}
      if(!closingFromHost && HasPendingChanges){DialogResult choice=MessageBox.Show(this,"Há alterações não aplicadas nem salvas em JSON. Salvar o projeto antes de fechar?","Mello Charts",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);if(choice==DialogResult.Cancel || (choice==DialogResult.Yes && !TrySaveProject()))e.Cancel=true;}
      base.OnFormClosing(e);
      if(!e.Cancel)acceptedClose=true;
    }
    private static void SideHeading(FlowLayoutPanel panel,string caption){panel.Controls.Add(new Label{Text=caption,AutoSize=true,Font=new Font("Segoe UI",9,FontStyle.Bold),ForeColor=Color.FromArgb(36,79,115),Margin=new Padding(0,8,0,12)});}
    private static void SideField(FlowLayoutPanel parent,string caption,Control input,int height){Panel panel=new Panel{Width=202,Height=height,Margin=new Padding(0,0,0,6)};Label label=new Label{Text=caption,Dock=DockStyle.Top,Height=20,ForeColor=Color.FromArgb(92,111,125)};input.Dock=DockStyle.Fill;panel.Controls.Add(input);panel.Controls.Add(label);parent.Controls.Add(panel);}
    private string SelectedPalette(){return palette.SelectedIndex==1?"forest":palette.SelectedIndex==2?"ink":"ocean";}
    private string SelectedStyle(){return style.SelectedIndex==1?"minimal":style.SelectedIndex==2?"dark":"executive";}
    private void PresentationTo(ChartDocument d){d.seriesActual=seriesActual.Text;d.seriesForecast=seriesForecast.Text;d.seriesBaseline=seriesBaseline.Text;d.fontFamily=Convert.ToString(fontFamily.SelectedItem)??"Arial";d.customColors=ColorTools.Clone(customColors);d.title=title.Text;d.subtitle=subtitle.Text;d.source=source.Text;d.unit=unit.Text;d.palette=SelectedPalette();d.style=SelectedStyle();d.compactNumbers=compact.Checked;d.showGrid=gridLines.Checked;d.decimals=Math.Max(0,decimals.SelectedIndex);d.showLabels=labels.Checked;d.showBaseline=baseline.Checked;d.includeTotal=total.Checked;}
    private void UpdateColorButton(){colorButton.Text=customColors==null?"Cores personalizadas…":"Cores personalizadas ✓";}
    private void CustomizeColors(){
      try{using(ChartColorForm form=new ChartColorForm(ReadEditor())){if(form.ShowDialog(this)==DialogResult.OK){customColors=ColorTools.Clone(form.Result);UpdateColorButton();RefreshChart();}}}
      catch(Exception ex){SetStatus(ex.Message,true);}
    }
    private static string TypeName(string type){return ChartCatalog.Get(type).Name;}
    private static void ConfigureTool(Button b,string text,Action action){b.Text=text;b.AutoSize=true;b.Height=31;b.Padding=new Padding(8,2,8,2);b.FlatStyle=FlatStyle.Flat;b.BackColor=Color.White;b.Margin=new Padding(0,3,8,0);b.FlatAppearance.BorderColor=Color.FromArgb(208,221,226);b.Click+=(s,e)=>action();}
    private static void Tool(FlowLayoutPanel p,string text,Action action){Button b=new Button();ConfigureTool(b,text,action);p.Controls.Add(b);}
    protected override bool ProcessCmdKey(ref Message msg,Keys key){if(key==(Keys.Control|Keys.Z)){UndoEdit();return true;}if(key==(Keys.Control|Keys.Y)){RedoEdit();return true;}if(key==(Keys.Control|Keys.S)){SaveProject();return true;}return base.ProcessCmdKey(ref msg,key);}
    private void UpdateHistoryButtons(){undoButton.Enabled=invalidEdit||undoStates.Count>0;redoButton.Enabled=redoStates.Count>0;refreshDataButton.Enabled=document.excelLink!=null;tips.SetToolTip(refreshDataButton,document.excelLink==null?"Vincule um arquivo XLSX salvo para atualizar seus dados.":document.excelLink.sheet+"!"+document.excelLink.range+" · "+document.excelLink.path+"\nLê a última versão salva; atualiza somente este gráfico.");}
    private void RestoreHistory(string state){historyRestoring=true;try{document=ChartCore.Deserialize(state);historyCurrent=state;invalidEdit=false;LoadDocument();}finally{historyRestoring=false;UpdateHistoryButtons();}}
    private void UndoEdit(){if(invalidEdit){RestoreHistory(historyCurrent);return;}if(undoStates.Count==0)return;redoStates.Add(historyCurrent);string state=undoStates[undoStates.Count-1];undoStates.RemoveAt(undoStates.Count-1);RestoreHistory(state);}
    private void RedoEdit(){if(redoStates.Count==0)return;undoStates.Add(historyCurrent);string state=redoStates[redoStates.Count-1];redoStates.RemoveAt(redoStates.Count-1);RestoreHistory(state);}
    private void LinkExcel(){ChooseExcel(null);}
    private void ChooseExcel(string path){try{using(WorkbookImportForm f=new WorkbookImportForm(document.type)){if(path!=null)f.SetPath(path);if(f.ShowDialog(this)==DialogResult.OK)ImportData(f.Result,"Dados vinculados ao Excel");}}catch(Exception ex){SetStatus(ex.Message,true);}}
    private void RefreshExcel(){try{ChartDocument refreshed=WorkbookData.Refresh(ReadEditor());document=refreshed;LoadDocument();SetStatus("Dados atualizados da versão salva do Excel. Aplique ao slide para concluir.",false);}catch(Exception ex){SetStatus("Dados anteriores preservados. "+ex.Message,true);}}

    private void SetStatus(string text,bool error){status.Text=text;status.ForeColor=error?Color.FromArgb(164,73,48):Color.FromArgb(67,97,111);}
    private void LoadDocument(){
      loading=true;seriesActual.Text=document.seriesActual;seriesForecast.Text=document.seriesForecast;seriesBaseline.Text=document.seriesBaseline;fontFamily.SelectedItem=document.fontFamily;
      bool comparative=!new[]{"waterfall","gantt","scatter","bubble"}.Contains(document.type)&&!ChartCatalog.Get(document.type).SingleValue;foreach(TextBox box in new[]{seriesActual,seriesForecast,seriesBaseline})box.Parent.Visible=comparative;
      customColors=ColorTools.Clone(document.customColors);UpdateColorButton();ChartDefinition definition=ChartCatalog.Get(document.type);typeLabel.Text=TypeName(document.type).ToUpperInvariant();title.Text=document.title;title.MaxLength=80;subtitle.Text=document.subtitle;subtitle.MaxLength=110;source.Text=document.source;source.MaxLength=100;unit.Text=document.unit;unit.MaxLength=16;palette.SelectedIndex=document.palette=="forest"?1:document.palette=="ink"?2:0;decimals.SelectedIndex=document.decimals;style.SelectedIndex=document.style=="minimal"?1:document.style=="dark"?2:0;compact.Checked=document.compactNumbers;gridLines.Checked=document.showGrid;labels.Checked=document.showLabels;baseline.Checked=document.showBaseline;total.Checked=document.includeTotal;baseline.Visible=definition.Baseline;total.Visible=definition.Total;gridLines.Visible=document.type!="pie"&&document.type!="donut";
      grid.Columns.Clear();grid.Rows.Clear();grid.Columns.Add("label","Categoria");grid.Columns[0].FillWeight=180;
      if(document.type=="gantt"){
        grid.Columns.Add("startDate","Início · aaaa-mm-dd");grid.Columns.Add("endDate","Fim · aaaa-mm-dd");grid.Columns.Add("progress","Progresso (%)");foreach(ChartRow r in document.rows)grid.Rows.Add(r.label,r.startDate,r.endDate,r.progress.ToString("R",Pt));
      }else if(document.type=="scatter"||document.type=="bubble"){
        grid.Columns.Add("x","Eixo X");grid.Columns.Add("y","Eixo Y");if(document.type=="bubble")grid.Columns.Add("size","Tamanho (área)");foreach(ChartRow r in document.rows){if(document.type=="bubble")grid.Rows.Add(r.label,r.x.ToString("R",Pt),r.y.ToString("R",Pt),r.size.ToString("R",Pt));else grid.Rows.Add(r.label,r.x.ToString("R",Pt),r.y.ToString("R",Pt));}
      }else if(document.type=="waterfall"){
        grid.Columns.Add("value","Valor");DataGridViewComboBoxColumn kind=new DataGridViewComboBoxColumn{Name="kind",HeaderText="Tipo",DataSource=new[]{"base","delta","total"}};grid.Columns.Add(kind);
        foreach(ChartRow r in document.rows)grid.Rows.Add(r.label,r.value.ToString("R",Pt),r.kind);
        grid.Rows[0].Cells["kind"].ReadOnly=true;
      }else if(definition.SingleValue){
        grid.Columns.Add("value","Valor");foreach(ChartRow r in document.rows)grid.Rows.Add(r.label,r.value.ToString("R",Pt));
      }else{
        grid.Columns.Add("actual",document.seriesActual);grid.Columns.Add("forecast",document.type=="buildup"&&document.seriesForecast=="Projetado"?"Projetado restante":document.seriesForecast);grid.Columns.Add("baseline",document.seriesBaseline);foreach(ChartRow r in document.rows)grid.Rows.Add(r.label,r.actual.ToString("R",Pt),r.forecast.ToString("R",Pt),r.baseline.ToString("R",Pt));
      }
      method.Text=definition.Method;tips.SetToolTip(method,definition.Method+"\nCabeçalho: "+definition.Header);tips.SetToolTip(grid,"Cole ou importe uma tabela com o cabeçalho: "+definition.Header);loading=false;RefreshChart();
    }
    private static double Number(object value){
      return ChartCore.ParseNumber(Convert.ToString(value,Pt));
    }
    private ChartDocument ReadEditor(){
      ChartDocument d=ChartCore.Clone(document);PresentationTo(d);d.rows.Clear();
      foreach(DataGridViewRow row in grid.Rows){ChartRow r=new ChartRow{label=Convert.ToString(row.Cells["label"].Value)};if(d.type=="gantt"){r.startDate=Convert.ToString(row.Cells["startDate"].Value);r.endDate=Convert.ToString(row.Cells["endDate"].Value);r.progress=Number(row.Cells["progress"].Value);}else if(d.type=="scatter"||d.type=="bubble"){r.x=Number(row.Cells["x"].Value);r.y=Number(row.Cells["y"].Value);if(d.type=="bubble")r.size=Number(row.Cells["size"].Value);}else if(d.type=="waterfall"){r.kind=Convert.ToString(row.Cells["kind"].Value);r.value=r.kind=="total"?0:Number(row.Cells["value"].Value);}else if(ChartCatalog.Get(d.type).SingleValue){r.value=Number(row.Cells["value"].Value);}else{r.actual=Number(row.Cells["actual"].Value);r.forecast=Number(row.Cells["forecast"].Value);r.baseline=Number(row.Cells["baseline"].Value);}d.rows.Add(r);}ChartCore.Validate(d);return d;
    }
    private void RefreshChart(){if(loading)return;try{ChartDocument d=ReadEditor();Scene nextScene=ChartCore.BuildScene(d);string serialized=ChartCore.Serialize(d);
      if(!historyRestoring && historyCurrent!=serialized){undoStates.Add(historyCurrent);if(undoStates.Count>80)undoStates.RemoveAt(0);redoStates.Clear();historyCurrent=serialized;}invalidEdit=false;
      preview.Scene=nextScene;lastValidDocument=ChartCore.Clone(d);snapshotDirty=serialized!=savedState;valid=true;apply.Enabled=true;SetStatus((context==null?"Prévia atualizada":"Destino: slide "+context.SlideNumber)+" · "+d.rows.Count+" categorias · Unidade: "+d.unit,false);}catch(Exception ex){invalidEdit=true;snapshotDirty=true;valid=false;apply.Enabled=false;SetStatus("Prévia do último estado válido · "+ex.Message,true);}UpdateHistoryButtons();preview.Invalidate();}
    private void AddRow(){if(grid.Rows.Count>=ChartCore.RowLimit(document.type)){SetStatus("Use até "+ChartCore.RowLimit(document.type)+" linhas neste gráfico.",true);return;}if(document.type=="gantt"){grid.EndEdit();ChartDocument current;try{current=ReadEditor();}catch(Exception ex){SetStatus(ex.Message,true);return;}string start=current.rows.Last().endDate;DateTime begin=DateTime.ParseExact(start,"yyyy-MM-dd",CultureInfo.InvariantCulture),end=begin.AddDays(Math.Min(7,(DateTime.MaxValue.Date-begin).TotalDays));grid.Rows.Add("Nova tarefa",start,end.ToString("yyyy-MM-dd"),"0");}else if(document.type=="bubble")grid.Rows.Add("Novo ponto","0","0","1");else if(document.type=="scatter")grid.Rows.Add("Novo ponto","0","0");else if(document.type=="waterfall"){int index=grid.Rows.Count;if(index>1&&Convert.ToString(grid.Rows[index-1].Cells["kind"].Value)=="total")index--;grid.Rows.Insert(index,new object[]{"Nova etapa","0","delta"});}else if(ChartCatalog.Get(document.type).SingleValue)grid.Rows.Add("Nova categoria","0");else grid.Rows.Add("Nova categoria","0","0","0");RefreshChart();}
    private void DeleteRow(){if(grid.CurrentCell==null)return;int index=grid.CurrentCell.RowIndex;if(grid.Rows.Count<=1||(document.type=="waterfall"&&index==0)){SetStatus("Mantenha a base inicial e pelo menos uma categoria.",true);return;}grid.Rows.RemoveAt(index);RefreshChart();}
    private void Import(){
      using(OpenFileDialog f=new OpenFileDialog{Filter="Tabelas CSV, TSV ou Excel|*.csv;*.tsv;*.txt;*.xlsx",Title="Importar tabela do gráfico"}){
        if(f.ShowDialog()!=DialogResult.OK)return;if(Path.GetExtension(f.FileName).Equals(".xlsx",StringComparison.OrdinalIgnoreCase)){ChooseExcel(f.FileName);return;}try{ChartDocument next=ChartCore.ImportDelimited(ReadLimited(f.FileName,1000000),document.type);
          ImportData(next,"Dados importados pelo usuário");
        }catch(Exception ex){SetStatus(ex.Message,true);}
      }
    }
    private void PasteTable(){
      try{if(!Clipboard.ContainsText())throw new InvalidOperationException("Copie uma tabela incluindo o cabeçalho: "+ChartCatalog.Get(document.type).Header+".");
        ChartDocument next=ChartCore.ImportDelimited(Clipboard.GetText(),document.type);ImportData(next,"Dados colados pelo usuário");
      }catch(Exception ex){SetStatus(ex.Message,true);}
    }
    private void ImportData(ChartDocument next,string importedSource){PresentationTo(next);next.source=importedSource;ChartCore.Validate(next);document=next;LoadDocument();}
    private static string ReadLimited(string file,long limit){if(new FileInfo(file).Length>limit)throw new InvalidOperationException("O arquivo é maior que o limite permitido.");return File.ReadAllText(file,EncodingUtf8);}
    private static readonly System.Text.Encoding EncodingUtf8=new System.Text.UTF8Encoding(false,true);
    private void OpenProject(){OpenProjectAt(null);}
    private void OpenRecovery(){if(!Directory.Exists(RecoveryStore.Folder)){SetStatus("Ainda não há rascunhos recuperados neste usuário do Windows.",false);return;}OpenProjectAt(RecoveryStore.Folder);}
    private void OpenProjectAt(string folder){using(OpenFileDialog f=new OpenFileDialog{Filter="Projeto Mello Charts|*.json",Title="Abrir projeto Mello Charts"}){if(folder!=null)f.InitialDirectory=folder;if(f.ShowDialog(this)!=DialogResult.OK)return;try{ChartDocument next=ChartCore.Deserialize(ReadLimited(f.FileName,1000000));if((target!=null||HasPendingChanges)&&MessageBox.Show(this,"Os dados do projeto substituirão os dados atuais deste editor. Continuar?","Abrir projeto",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;document=next;LoadDocument();}catch(Exception ex){SetStatus(ex.Message,true);}}}
    private void SaveProject(){TrySaveProject();}
    private bool TrySaveProject(){try{grid.EndEdit();ChartDocument d=ReadEditor();using(SaveFileDialog f=new SaveFileDialog{Filter="Projeto Mello Charts|*.json",FileName="mello-charts.json"}){if(f.ShowDialog(this)!=DialogResult.OK)return false;File.WriteAllText(f.FileName,ChartCore.Serialize(d),EncodingUtf8);savedState=ChartCore.Serialize(d);snapshotDirty=false;SetStatus("Projeto salvo. Para aplicar ao slide, use o botão de inserir ou atualizar.",false);return true;}}catch(Exception ex){SetStatus(ex.Message,true);return false;}}
    private void ExportPng(){
      try{grid.EndEdit();ChartDocument d=ReadEditor();using(SaveFileDialog f=new SaveFileDialog{Filter="Imagem PNG|*.png",FileName="mello-charts.png"}){if(f.ShowDialog()!=DialogResult.OK)return;using(Bitmap bitmap=ChartPreview.RenderImage(ChartCore.BuildScene(d),1920,1080))bitmap.Save(f.FileName,System.Drawing.Imaging.ImageFormat.Png);SetStatus("PNG salvo em 1920 × 1080. Para editar no PowerPoint, use Inserir no slide.",false);}}catch(Exception ex){SetStatus(ex.Message,true);}
    }
    private void ExportCsv(){try{ChartDocument d=ReadEditor();using(SaveFileDialog f=new SaveFileDialog{Filter="Tabela CSV|*.csv",FileName="mello-charts.csv"}){if(f.ShowDialog()==DialogResult.OK)File.WriteAllText(f.FileName,ChartCore.ToCsv(d),new System.Text.UTF8Encoding(true));}}catch(Exception ex){SetStatus(ex.Message,true);}}
    private void ApplyChart(){
      grid.EndEdit();RefreshChart();if(!valid)return;try{ChartDocument d=ReadEditor();apply.Enabled=false;
        if(application==null){object ppt;try{ppt=Marshal.GetActiveObject("PowerPoint.Application");}catch(COMException){ppt=Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application",true));}dynamic app=ppt;app.Visible=-1;dynamic presentation=app.Presentations.Add(-1);presentation.PageSetup.SlideWidth=960;presentation.PageSetup.SlideHeight=540;dynamic slide=presentation.Slides.Add(1,12);app.ActiveWindow.View.GotoSlide(1);application=ppt;context=PowerPointBridge.CaptureContext(application);Text="Mello Charts · nova apresentação · slide 1";}
        target=PowerPointBridge.InsertInContext(application,d,target,context);document=d;savedState=ChartCore.Serialize(d);snapshotDirty=false;apply.Text="Atualizar gráfico";SetStatus("Gráfico pronto no slide de origem. Salve a apresentação. Você pode continuar editando aqui.",false);}catch(Exception ex){SetStatus(ex.Message,true);MessageBox.Show(this,ex.Message,"Mello Charts",MessageBoxButtons.OK,MessageBoxIcon.Warning);}finally{apply.Enabled=valid;}
    }
  }
  public sealed class ChartPreview:Control {
    public Scene Scene;
    public ChartPreview(){DoubleBuffered=true;BackColor=Color.FromArgb(231,237,242);ResizeRedraw=true;}
    public static Bitmap RenderImage(Scene scene,int width,int height){
      if(scene==null)throw new ArgumentNullException("scene");if(width<1 || height<1 || width>4096 || height>4096)throw new ArgumentException("Use dimensões de 1 a 4096 pixels.");
      Bitmap bitmap=new Bitmap(width,height,System.Drawing.Imaging.PixelFormat.Format32bppArgb);
      try{using(Graphics graphics=Graphics.FromImage(bitmap)){graphics.Clear(Color.White);DrawScene(graphics,scene,new RectangleF(0,0,width,height));}return bitmap;}catch{bitmap.Dispose();throw;}
    }
    protected override void OnPaint(PaintEventArgs e){
      base.OnPaint(e);if(Scene==null){using(Brush brush=new SolidBrush(Color.FromArgb(110,130,144)))e.Graphics.DrawString("Revise os dados para visualizar o gráfico.",Font,brush,20,20);return;}
      DrawScene(e.Graphics,Scene,ClientRectangle);
    }
    private static void DrawScene(Graphics graphics,Scene scene,RectangleF bounds){
      if(bounds.Width<=0 || bounds.Height<=0)return;
      float scale=Math.Min(bounds.Width/(float)scene.Width,bounds.Height/(float)scene.Height),ox=bounds.X+(bounds.Width-(float)scene.Width*scale)/2,oy=bounds.Y+(bounds.Height-(float)scene.Height*scale)/2;
      System.Drawing.Drawing2D.GraphicsState state=graphics.Save();
      try{
        graphics.TranslateTransform(ox,oy);graphics.ScaleTransform(scale,scale);graphics.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;graphics.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        foreach(SceneElement s in scene.Elements){Color color=ColorTranslator.FromHtml(s.Color);
          if(s.Type=="rect"){using(Brush brush=new SolidBrush(color))graphics.FillRectangle(brush,(float)s.X,(float)s.Y,(float)s.W,(float)s.H);}
          else if(s.Type=="ellipse"){using(Brush brush=new SolidBrush(color))graphics.FillEllipse(brush,(float)s.X,(float)s.Y,(float)s.W,(float)s.H);}
          else if(s.Type=="polygon"){
            if(s.Points==null||s.Points.Length<6||s.Points.Length%2!=0)throw new InvalidOperationException("Polígono inválido na cena do gráfico.");
            PointF[] points=new PointF[s.Points.Length/2];for(int i=0;i<points.Length;i++)points[i]=new PointF((float)s.Points[i*2],(float)s.Points[i*2+1]);
            using(Brush brush=new SolidBrush(color))graphics.FillPolygon(brush,points);
          }
          else if(s.Type=="line"){using(Pen pen=new Pen(color,(float)Math.Max(.5,s.Size))){if(s.Dash)pen.DashStyle=System.Drawing.Drawing2D.DashStyle.Dash;graphics.DrawLine(pen,(float)s.X,(float)s.Y,(float)s.X2,(float)s.Y2);}}
          else{using(Font font=new Font(s.FontFamily??"Arial",(float)s.Size,s.Bold?FontStyle.Bold:FontStyle.Regular,GraphicsUnit.Pixel))using(Brush brush=new SolidBrush(color))using(StringFormat f=new StringFormat()){f.Alignment=s.Align=="center"?StringAlignment.Center:s.Align=="right"?StringAlignment.Far:StringAlignment.Near;f.LineAlignment=StringAlignment.Near;graphics.DrawString(s.Text,font,brush,new RectangleF((float)s.X,(float)s.Y,(float)s.W,(float)s.H),f);}}
        }
      }finally{graphics.Restore(state);}
    }
  }
}
