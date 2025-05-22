namespace Slide2Pdf
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.btnExportFullSlide = this.Factory.CreateRibbonButton();
            this.btnExportContent = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnExportFullSlide);
            this.group1.Items.Add(this.btnExportContent);
            this.group1.Label = "Export PDF";
            this.group1.Name = "group1";
            // 
            // btnExportFullSlide
            // 
            this.btnExportFullSlide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportFullSlide.Image = global::Slide2Pdf.Properties.Resources.document_pdf_512x512;
            this.btnExportFullSlide.Label = "Full Slide";
            this.btnExportFullSlide.Name = "btnExportFullSlide";
            this.btnExportFullSlide.ScreenTip = "Export current slide as PDF";
            this.btnExportFullSlide.ShowImage = true;
            this.btnExportFullSlide.SuperTip = "It remembers the output file for each slide. To export to another location, hold " +
    "\"shift\" while clicking this button.";
            this.btnExportFullSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportSlideToPdf_Click);
            // 
            // btnExportContent
            // 
            this.btnExportContent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExportContent.Image = global::Slide2Pdf.Properties.Resources.crop_512x512;
            this.btnExportContent.Label = "Cropped Slide";
            this.btnExportContent.Name = "btnExportContent";
            this.btnExportContent.ScreenTip = "Export current slide as PDF, cropping to visible content";
            this.btnExportContent.ShowImage = true;
            this.btnExportContent.SuperTip = "It remembers the output file for each slide. To export to another location, hold " +
    "\"shift\" while clicking this button.";
            this.btnExportContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportContent_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportFullSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportContent;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
