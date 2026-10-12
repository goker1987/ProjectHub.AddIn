namespace ProjectHub.AddIn
{
    partial class ProjectHubRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        ///Gerekli tasarımcı değişkeni.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ProjectHubRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        ///Kullanılan tüm kaynakları temizleyin.
        /// </summary>
        /// <param name="disposing">yönetilen kaynaklar dispose edilmeliyse doğru; aksi halde yanlış.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Bileşen Tasarımcısı üretimi kod

        /// <summary>
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grpProject = this.Factory.CreateRibbonGroup();
            this.btnConnect = this.Factory.CreateRibbonButton();
            this.btnPublish = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpProject.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.grpProject);
            this.tab1.Label = "Project Operations";
            this.tab1.Name = "tab1";
            // 
            // grpProject
            // 
            this.grpProject.Items.Add(this.btnConnect);
            this.grpProject.Items.Add(this.btnPublish);
            this.grpProject.Label = "Build 1.0.1.0";
            this.grpProject.Name = "grpProject";
            // 
            // btnConnect
            // 
            this.btnConnect.Label = "Connect";
            this.btnConnect.Name = "btnConnect";
            this.btnConnect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConnect_Click);
            // 
            // btnPublish
            // 
            this.btnPublish.Enabled = false;
            this.btnPublish.Label = "Publish";
            this.btnPublish.Name = "btnPublish";
            this.btnPublish.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPublish_Click);
            // 
            // ProjectHubRibbon
            // 
            this.Name = "ProjectHubRibbon";
            this.RibbonType = "Microsoft.Project.Project";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ProjectHubRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpProject.ResumeLayout(false);
            this.grpProject.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpProject;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConnect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPublish;
    }

    partial class ThisRibbonCollection
    {
        internal ProjectHubRibbon ProjectHubRibbon
        {
            get { return this.GetRibbon<ProjectHubRibbon>(); }
        }
    }
}
