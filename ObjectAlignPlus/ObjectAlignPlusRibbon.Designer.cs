namespace ObjectAlignPlus
{
    partial class ObjectAlignPlusRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ObjectAlignPlusRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonBox boxSpacing;
            Microsoft.Office.Tools.Ribbon.RibbonBox boxAlignH;
            Microsoft.Office.Tools.Ribbon.RibbonBox boxAlignV;
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpObjectAlignPlus = this.Factory.CreateRibbonGroup();
            this.ebSpacing = this.Factory.CreateRibbonEditBox();
            this.btVert = this.Factory.CreateRibbonButton();
            this.btVertBottomAlign = this.Factory.CreateRibbonButton();
            this.btSpPlus = this.Factory.CreateRibbonButton();
            this.btSpMinus = this.Factory.CreateRibbonButton();
            this.btHorz = this.Factory.CreateRibbonButton();
            this.btHorzRightAlign = this.Factory.CreateRibbonButton();
            boxSpacing = this.Factory.CreateRibbonBox();
            boxAlignH = this.Factory.CreateRibbonBox();
            boxAlignV = this.Factory.CreateRibbonBox();
            this.tab1.SuspendLayout();
            this.grpObjectAlignPlus.SuspendLayout();
            boxSpacing.SuspendLayout();
            boxAlignH.SuspendLayout();
            boxAlignV.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpObjectAlignPlus);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grpObjectAlignPlus
            // 
            this.grpObjectAlignPlus.Items.Add(boxSpacing);
            this.grpObjectAlignPlus.Items.Add(boxAlignH);
            this.grpObjectAlignPlus.Items.Add(boxAlignV);
            this.grpObjectAlignPlus.Label = "ObjectAlign+";
            this.grpObjectAlignPlus.Name = "grpObjectAlignPlus";
            // 
            // boxSpacing
            // 
            boxSpacing.Items.Add(this.ebSpacing);
            boxSpacing.Items.Add(this.btSpPlus);
            boxSpacing.Items.Add(this.btSpMinus);
            boxSpacing.Name = "boxSpacing";
            // 
            // ebSpacing
            // 
            this.ebSpacing.Label = "間隔";
            this.ebSpacing.MaxLength = 3;
            this.ebSpacing.Name = "ebSpacing";
            this.ebSpacing.Text = "0";
            // 
            // boxAlignH
            // 
            boxAlignH.Items.Add(this.btHorz);
            boxAlignH.Items.Add(this.btHorzRightAlign);
            boxAlignH.Name = "boxAlignH";
            // 
            // btVert
            // 
            this.btVert.Image = global::ObjectAlignPlus.Properties.Resources.space_v;
            this.btVert.Label = "垂直分布：上から";
            this.btVert.Name = "btVert";
            this.btVert.ShowImage = true;
            this.btVert.SuperTip = "上端基準で垂直方向等間隔に分布";
            this.btVert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btVert_Click);
            // 
            // btVertBottomAlign
            // 
            this.btVertBottomAlign.Image = global::ObjectAlignPlus.Properties.Resources.arrow_up;
            this.btVertBottomAlign.Label = "下から";
            this.btVertBottomAlign.Name = "btVertBottomAlign";
            this.btVertBottomAlign.ShowImage = true;
            this.btVertBottomAlign.SuperTip = "下端基準で垂直方向等間隔に分布";
            this.btVertBottomAlign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btVertBottomAlign_Click);
            // 
            // btSpPlus
            // 
            this.btSpPlus.Image = global::ObjectAlignPlus.Properties.Resources.action_add_16xMD;
            this.btSpPlus.Label = "+";
            this.btSpPlus.Name = "btSpPlus";
            this.btSpPlus.ShowImage = true;
            this.btSpPlus.ShowLabel = false;
            this.btSpPlus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btSpPlus_Click);
            // 
            // btSpMinus
            // 
            this.btSpMinus.Image = global::ObjectAlignPlus.Properties.Resources.Symbols_Blocked_16xLG;
            this.btSpMinus.Label = "-";
            this.btSpMinus.Name = "btSpMinus";
            this.btSpMinus.ShowImage = true;
            this.btSpMinus.ShowLabel = false;
            this.btSpMinus.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btSpMinus_Click);
            // 
            // btHorz
            // 
            this.btHorz.Image = global::ObjectAlignPlus.Properties.Resources.space_h;
            this.btHorz.Label = "水平分布：左から";
            this.btHorz.Name = "btHorz";
            this.btHorz.ShowImage = true;
            this.btHorz.SuperTip = "左端基準で水平方向等間隔に分布";
            this.btHorz.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btHorz_Click);
            // 
            // btHorzRightAlign
            // 
            this.btHorzRightAlign.Image = global::ObjectAlignPlus.Properties.Resources.arrow_left;
            this.btHorzRightAlign.Label = "右から";
            this.btHorzRightAlign.Name = "btHorzRightAlign";
            this.btHorzRightAlign.ShowImage = true;
            this.btHorzRightAlign.SuperTip = "右端基準で水平方向等間隔に分布";
            this.btHorzRightAlign.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btHorzRightAlign_Click);
            // 
            // boxAlignV
            // 
            boxAlignV.Items.Add(this.btVert);
            boxAlignV.Items.Add(this.btVertBottomAlign);
            boxAlignV.Name = "boxAlignV";
            // 
            // ObjectAlignPlusRibbon
            // 
            this.Name = "ObjectAlignPlusRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpObjectAlignPlus.ResumeLayout(false);
            this.grpObjectAlignPlus.PerformLayout();
            boxSpacing.ResumeLayout(false);
            boxSpacing.PerformLayout();
            boxAlignH.ResumeLayout(false);
            boxAlignH.PerformLayout();
            boxAlignV.ResumeLayout(false);
            boxAlignV.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpObjectAlignPlus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btHorz;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btVert;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox ebSpacing;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btSpPlus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btSpMinus;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btHorzRightAlign;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btVertBottomAlign;
    }

    partial class ThisRibbonCollection
    {
        internal ObjectAlignPlusRibbon Ribbon1
        {
            get { return this.GetRibbon<ObjectAlignPlusRibbon>(); }
        }
    }
}
