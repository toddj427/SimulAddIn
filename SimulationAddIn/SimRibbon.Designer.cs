namespace SimulationAddIn
{
    partial class SimRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public SimRibbon()
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
            this.simTab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ModelControl = this.Factory.CreateRibbonGroup();
            this.grpHelp = this.Factory.CreateRibbonGroup();
            this.btnInitialize = this.Factory.CreateRibbonButton();
            this.btnModelPath = this.Factory.CreateRibbonButton();
            this.BtnModelName = this.Factory.CreateRibbonButton();
            this.BtnRunModel = this.Factory.CreateRibbonButton();
            this.BtnRunWindowless = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.BtnHelp = this.Factory.CreateRibbonButton();
            this.simTab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.ModelControl.SuspendLayout();
            this.grpHelp.SuspendLayout();
            this.SuspendLayout();
            // 
            // simTab1
            // 
            this.simTab1.Groups.Add(this.group1);
            this.simTab1.Groups.Add(this.ModelControl);
            this.simTab1.Groups.Add(this.grpHelp);
            this.simTab1.Label = "Simulation";
            this.simTab1.Name = "simTab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnInitialize);
            this.group1.Items.Add(this.btnModelPath);
            this.group1.Items.Add(this.BtnModelName);
            this.group1.Label = "Set Up";
            this.group1.Name = "group1";
            // 
            // ModelControl
            // 
            this.ModelControl.Items.Add(this.BtnRunModel);
            this.ModelControl.Items.Add(this.BtnRunWindowless);
            this.ModelControl.Label = "(Not Implemented)";
            this.ModelControl.Name = "ModelControl";
            // 
            // grpHelp
            // 
            this.grpHelp.Items.Add(this.button1);
            this.grpHelp.Items.Add(this.BtnHelp);
            this.grpHelp.Name = "grpHelp";
            // 
            // btnInitialize
            // 
            this.btnInitialize.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInitialize.Description = "Initialize a New UI";
            this.btnInitialize.Image = global::SimulationAddIn.Properties.Resources.Start;
            this.btnInitialize.Label = "Initialize";
            this.btnInitialize.Name = "btnInitialize";
            this.btnInitialize.ShowImage = true;
            this.btnInitialize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnInitialize_Click);
            // 
            // btnModelPath
            // 
            this.btnModelPath.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModelPath.Image = global::SimulationAddIn.Properties.Resources.foldericon;
            this.btnModelPath.Label = "Model Path";
            this.btnModelPath.Name = "btnModelPath";
            this.btnModelPath.ScreenTip = "Set Model Path";
            this.btnModelPath.ShowImage = true;
            this.btnModelPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModelPath_Click);
            // 
            // BtnModelName
            // 
            this.BtnModelName.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnModelName.Description = "Model Name";
            this.BtnModelName.Image = global::SimulationAddIn.Properties.Resources.File;
            this.BtnModelName.Label = "Model Name";
            this.BtnModelName.Name = "BtnModelName";
            this.BtnModelName.ShowImage = true;
            this.BtnModelName.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnModelName_Click);
            // 
            // BtnRunModel
            // 
            this.BtnRunModel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnRunModel.Image = global::SimulationAddIn.Properties.Resources.Run;
            this.BtnRunModel.Label = "Run Model";
            this.BtnRunModel.Name = "BtnRunModel";
            this.BtnRunModel.ShowImage = true;
            this.BtnRunModel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRunmodel_Click);
            // 
            // BtnRunWindowless
            // 
            this.BtnRunWindowless.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnRunWindowless.Image = global::SimulationAddIn.Properties.Resources.Run;
            this.BtnRunWindowless.Label = "Run w/o Animation";
            this.BtnRunWindowless.Name = "BtnRunWindowless";
            this.BtnRunWindowless.ShowImage = true;
            this.BtnRunWindowless.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnRunWindowless_Click);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::SimulationAddIn.Properties.Resources.SheetIcon;
            this.button1.Label = "SheetForm";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSheetSelector_Click);
            // 
            // BtnHelp
            // 
            this.BtnHelp.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnHelp.Image = global::SimulationAddIn.Properties.Resources.helpicon;
            this.BtnHelp.Label = "Help";
            this.BtnHelp.Name = "BtnHelp";
            this.BtnHelp.ShowImage = true;
            this.BtnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnHelp_Click);
            // 
            // SimRibbon
            // 
            this.Name = "SimRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.simTab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SimRibbon_Load);
            this.simTab1.ResumeLayout(false);
            this.simTab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ModelControl.ResumeLayout(false);
            this.ModelControl.PerformLayout();
            this.grpHelp.ResumeLayout(false);
            this.grpHelp.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab simTab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelPath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInitialize;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ModelControl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRunModel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnModelName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnRunWindowless;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal SimRibbon SimRibbon
        {
            get { return this.GetRibbon<SimRibbon>(); }
        }
    }
}
