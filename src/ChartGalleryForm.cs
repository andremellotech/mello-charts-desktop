using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace MelloCharts {
  // The gallery only chooses a model. The host opens a separate editor afterwards.
  public sealed class ChartGalleryForm:Form {
    private readonly FlowLayoutPanel cards=new FlowLayoutPanel();
    private readonly List<ChartCard> models=new List<ChartCard>();
    private readonly Label selection=new Label(),purpose=new Label();
    private readonly Button choose=new Button();
    public string SelectedType {get;private set;}

    public ChartGalleryForm(){
      Text="Mello Charts · Galeria de modelos";Font=new Font("Segoe UI",9);BackColor=Color.FromArgb(239,243,246);
      AutoScaleMode=AutoScaleMode.Dpi;StartPosition=FormStartPosition.CenterParent;ShowInTaskbar=false;KeyPreview=true;
      Rectangle area=Screen.PrimaryScreen.WorkingArea;Size=new Size(Math.Min(1160,area.Width-32),Math.Min(900,area.Height-32));MinimumSize=new Size(Math.Min(740,area.Width-32),Math.Min(560,area.Height-32));
      TableLayoutPanel root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(20),Margin=new Padding(0)};
      root.RowStyles.Add(new RowStyle(SizeType.Absolute,78));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Absolute,104));Controls.Add(root);
      Panel heading=new Panel{Dock=DockStyle.Fill,Margin=new Padding(0)};
      heading.Controls.Add(new Label{Text="Escolha a melhor forma de contar sua história",AutoSize=false,Dock=DockStyle.Top,Height=34,Font=new Font("Segoe UI",17,FontStyle.Bold),ForeColor=Color.FromArgb(28,62,80)});
      heading.Controls.Add(new Label{Text="11 modelos editáveis · selecione para ver a finalidade e use o modelo para começar",Dock=DockStyle.Bottom,Height=32,ForeColor=Color.FromArgb(92,111,125)});root.Controls.Add(heading,0,0);
      cards.Dock=DockStyle.Fill;cards.AutoScroll=true;cards.WrapContents=true;cards.Padding=new Padding(0,0,0,8);cards.Margin=new Padding(0);root.Controls.Add(cards,0,1);
      TableLayoutPanel footer=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=2,RowCount=1,Margin=new Padding(0,12,0,0)};footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute,170));root.Controls.Add(footer,0,2);
      Panel detail=new Panel{Dock=DockStyle.Fill,Margin=new Padding(0,0,14,0)};selection.Dock=DockStyle.Top;selection.Height=26;selection.Font=new Font("Segoe UI",11,FontStyle.Bold);selection.ForeColor=Color.FromArgb(28,62,80);purpose.Dock=DockStyle.Fill;purpose.ForeColor=Color.FromArgb(70,95,110);purpose.Padding=new Padding(0,4,0,0);detail.Controls.Add(purpose);detail.Controls.Add(selection);footer.Controls.Add(detail,0,0);
      TableLayoutPanel actions=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Margin=new Padding(0)};actions.RowStyles.Add(new RowStyle(SizeType.Percent,55));actions.RowStyles.Add(new RowStyle(SizeType.Percent,45));footer.Controls.Add(actions,1,0);
      choose.Text="Usar modelo";choose.Dock=DockStyle.Fill;choose.FlatStyle=FlatStyle.Flat;choose.FlatAppearance.BorderSize=0;choose.BackColor=Color.FromArgb(36,79,115);choose.ForeColor=Color.White;choose.Click+=(s,e)=>Choose();actions.Controls.Add(choose,0,0);
      Button cancel=new Button{Text="Cancelar",Dock=DockStyle.Fill,FlatStyle=FlatStyle.Flat,BackColor=Color.White,DialogResult=DialogResult.Cancel};cancel.FlatAppearance.BorderColor=Color.FromArgb(208,221,226);actions.Controls.Add(cancel,0,1);AcceptButton=choose;CancelButton=cancel;
      try{
        foreach(ChartDefinition definition in ChartCatalog.All){
          ChartDocument example=ChartCore.Example(definition.Key);
          ChartCard card=new ChartCard(definition,ChartPreview.RenderImage(ChartCore.BuildScene(example),480,270));
          models.Add(card);cards.Controls.Add(card);card.Click+=(s,e)=>SelectCard((ChartCard)s);card.DoubleClick+=(s,e)=>{SelectCard((ChartCard)s);Choose();};card.Enter+=(s,e)=>SelectCard((ChartCard)s);
          card.KeyDown+=CardKeyDown;
        }
      }catch{Dispose();throw;}
      SelectCard(models[0]);
    }
    private void SelectCard(ChartCard card){
      SelectedType=card.Definition.Key;selection.Text=card.Definition.Name+" · "+card.Definition.Category;purpose.Text=card.Definition.Description+"\r\n"+card.Definition.Method;
      foreach(ChartCard model in models){model.Selected=Object.ReferenceEquals(model,card);model.Invalidate();}
    }
    private void Choose(){if(ChartCatalog.Contains(SelectedType)){DialogResult=DialogResult.OK;Close();}}
    private void CardKeyDown(object sender,KeyEventArgs e){
      ChartCard current=(ChartCard)sender;int index=models.IndexOf(current),next=index;
      if(e.KeyCode==Keys.Right)next++;else if(e.KeyCode==Keys.Left)next--;else if(e.KeyCode==Keys.Down||e.KeyCode==Keys.Up){
        int columns=Math.Max(1,cards.ClientSize.Width/(current.Width+current.Margin.Horizontal));next+=e.KeyCode==Keys.Down?columns:-columns;
      }else if(e.KeyCode==Keys.Space){SelectCard(current);e.Handled=true;return;}else return;
      next=Math.Max(0,Math.Min(models.Count-1,next));models[next].Focus();cards.ScrollControlIntoView(models[next]);e.Handled=true;e.SuppressKeyPress=true;
    }
    protected override void OnShown(EventArgs e){base.OnShown(e);if(models.Count>0)models[0].Focus();}

    private sealed class ChartCard:Control {
      public readonly ChartDefinition Definition;private readonly Bitmap thumbnail;public bool Selected;
      public ChartCard(ChartDefinition definition,Bitmap image){
        Definition=definition;thumbnail=image;Size=new Size(250,220);Margin=new Padding(0,0,14,14);Cursor=Cursors.Hand;TabStop=true;AccessibleName=definition.Name;AccessibleDescription=definition.Description;AccessibleRole=AccessibleRole.RadioButton;
        SetStyle(ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer|ControlStyles.Selectable|ControlStyles.StandardDoubleClick,true);
      }
      protected override bool IsInputKey(Keys keyData){Keys key=keyData&Keys.KeyCode;return key==Keys.Left||key==Keys.Right||key==Keys.Up||key==Keys.Down||base.IsInputKey(keyData);}
      protected override void OnMouseDown(MouseEventArgs e){base.OnMouseDown(e);if(e.Button==MouseButtons.Left)Focus();}
      protected override void OnPaint(PaintEventArgs e){
        base.OnPaint(e);Graphics g=e.Graphics;float scale=Width/258f;int pad=(int)(9*scale),imageWidth=Width-pad*2,imageHeight=(int)(imageWidth*9f/16f);
        g.Clear(Color.White);g.DrawImage(thumbnail,new Rectangle(pad,pad,imageWidth,imageHeight));
        using(Pen border=new Pen(Selected?Color.FromArgb(27,126,121):Color.FromArgb(208,221,226),Selected?3*scale:scale))g.DrawRectangle(border,1,1,Width-3,Height-3);
        Rectangle nameBox=new Rectangle(pad+3,pad+imageHeight+(int)(10*scale),Width-pad*2-6,(int)(25*scale));
        using(Font nameFont=new Font(Font.FontFamily,10,FontStyle.Bold))TextRenderer.DrawText(g,Definition.Name,nameFont,nameBox,Color.FromArgb(28,62,80),TextFormatFlags.SingleLine|TextFormatFlags.EndEllipsis);
        Rectangle categoryBox=new Rectangle(nameBox.X,nameBox.Bottom+2,nameBox.Width,(int)(25*scale));TextRenderer.DrawText(g,Definition.Category+(Selected?" · Selecionado":""),Font,categoryBox,Selected?Color.FromArgb(27,126,121):Color.FromArgb(92,111,125),TextFormatFlags.SingleLine|TextFormatFlags.EndEllipsis);
        if(Focused)ControlPaint.DrawFocusRectangle(g,new Rectangle(5,5,Width-10,Height-10));
      }
      protected override void Dispose(bool disposing){if(disposing)thumbnail.Dispose();base.Dispose(disposing);}
    }
  }
}
