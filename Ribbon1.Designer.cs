
namespace WebScrape
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
            this.ribbon = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.CSV = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.JSON = this.Factory.CreateRibbonButton();
            this.Text = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.Attachments = this.Factory.CreateRibbonButton();
            this.jsonAtt = this.Factory.CreateRibbonButton();
            this.TextAtt = this.Factory.CreateRibbonButton();
            this.ribbon.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // ribbon
            // 
            this.ribbon.Groups.Add(this.group1);
            this.ribbon.Groups.Add(this.group2);
            this.ribbon.Label = "Mail Scrape";
            this.ribbon.Name = "ribbon";
            // 
            // group1
            // 
            this.group1.Items.Add(this.CSV);
            this.group1.Items.Add(this.JSON);
            this.group1.Items.Add(this.Text);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Files";
            this.group1.Name = "group1";
            // 
            // CSV
            // 
            this.CSV.Label = "CSV Metadata";
            this.CSV.Name = "CSV";
            this.CSV.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CSV_Click);
            // 
            // button1
            // 
            this.button1.Label = "";
            this.button1.Name = "button1";
            // 
            // JSON
            // 
            this.JSON.Label = "Multiple JSON";
            this.JSON.Name = "JSON";
            this.JSON.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.JSON_Click);
            // 
            // Text
            // 
            this.Text.Label = "Text (.txt)";
            this.Text.Name = "Text";
            this.Text.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Text_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.Attachments);
            this.group2.Items.Add(this.jsonAtt);
            this.group2.Items.Add(this.TextAtt);
            this.group2.Label = "Attachments";
            this.group2.Name = "group2";
            // 
            // Attachments
            // 
            this.Attachments.Label = "All Attachments";
            this.Attachments.Name = "Attachments";
            this.Attachments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Attachments_Click);
            // 
            // jsonAtt
            // 
            this.jsonAtt.Label = "JSON And Attachments";
            this.jsonAtt.Name = "jsonAtt";
            this.jsonAtt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.jsonAtt_Click);
            // 
            // TextAtt
            // 
            this.TextAtt.Label = "Text And Attachments";
            this.TextAtt.Name = "TextAtt";
            this.TextAtt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TextAtt_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.ribbon);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.ribbon.ResumeLayout(false);
            this.ribbon.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ribbon;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CSV;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton JSON;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Text;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Attachments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton jsonAtt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TextAtt;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
