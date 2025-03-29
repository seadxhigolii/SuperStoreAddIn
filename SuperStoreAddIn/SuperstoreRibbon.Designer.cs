namespace SuperStoreAddIn
{
    partial class SuperstoreRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SuperstoreRibbon()
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
            this.btnHighlightHighSales = this.Factory.CreateRibbonButton();
            this.btnShowSummary = this.Factory.CreateRibbonButton();
            this.btnClearHighlights = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnHighlightHighSales);
            this.group1.Items.Add(this.btnShowSummary);
            this.group1.Items.Add(this.btnClearHighlights);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnHighlightHighSales
            // 
            this.btnHighlightHighSales.Label = "Highlight High Sales";
            this.btnHighlightHighSales.Name = "btnHighlightHighSales";
            this.btnHighlightHighSales.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHighlightHighSales_Click);
            // 
            // btnShowSummary
            // 
            this.btnShowSummary.Label = "Show Summary";
            this.btnShowSummary.Name = "btnShowSummary";
            this.btnShowSummary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnShowSummary_Click);
            // 
            // btnClearHighlights
            // 
            this.btnClearHighlights.Label = "Clear Highlights";
            this.btnClearHighlights.Name = "btnClearHighlights";
            this.btnClearHighlights.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearHighlights_Click);
            // 
            // SuperstoreRibbon
            // 
            this.Name = "SuperstoreRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SuperstoreRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHighlightHighSales;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnShowSummary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearHighlights;
    }

    partial class ThisRibbonCollection
    {
        internal SuperstoreRibbon SuperstoreRibbon
        {
            get { return this.GetRibbon<SuperstoreRibbon>(); }
        }
    }
}
