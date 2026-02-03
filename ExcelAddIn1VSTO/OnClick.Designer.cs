namespace ExcelAddIn1
{
    partial class OnClick : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OnClick()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnExportJSON = this.Factory.CreateRibbonButton();
            this.btnExportCSV = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnExportJSON);
            this.group1.Items.Add(this.btnExportCSV);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnExportJSON
            // 
            this.btnExportJSON.Label = "Export JSON";
            this.btnExportJSON.Name = "btnExportJSON";
            this.btnExportJSON.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExportJSON_Click);
            // 
            // btnExportCSV
            // 
            this.btnExportCSV.Label = "Export CSV";
            this.btnExportCSV.Name = "btnExportCSV";
            this.btnExportCSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExportCSV_Click);
            // 
            // OnClick
            // 
            this.Name = "OnClick";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OnClick_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportJSON;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportCSV;
    }

    partial class ThisRibbonCollection
    {
        internal OnClick OnClick
        {
            get { return this.GetRibbon<OnClick>(); }
        }
    }
}
