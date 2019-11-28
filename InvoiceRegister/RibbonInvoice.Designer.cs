namespace InvoiceRegister
{
    partial class RibbonInvoice : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonInvoice()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonInvoice));
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.LnInvoices = this.Factory.CreateRibbonTab();
            this.grpInvoices = this.Factory.CreateRibbonGroup();
            this.btn_FillClasificators = this.Factory.CreateRibbonButton();
            this.btnLoadInvoiceList = this.Factory.CreateRibbonButton();
            this.btnSaveInvoice = this.Factory.CreateRibbonButton();
            this.btnNewFromThis = this.Factory.CreateRibbonButton();
            this.btnOpenInvoice = this.Factory.CreateRibbonButton();
            this.btnOpenTemplate = this.Factory.CreateRibbonButton();
            this.btnGenerateInvoice = this.Factory.CreateRibbonButton();
            this.grpClients = this.Factory.CreateRibbonGroup();
            this.btnNewClient = this.Factory.CreateRibbonButton();
            this.btnLoadClientList = this.Factory.CreateRibbonButton();
            this.btnUpdateClientInfo = this.Factory.CreateRibbonButton();
            this.btnValidatePostalCode = this.Factory.CreateRibbonButton();
            this.LnInvoices.SuspendLayout();
            this.grpInvoices.SuspendLayout();
            this.grpClients.SuspendLayout();
            this.SuspendLayout();
            // 
            // LnInvoices
            // 
            this.LnInvoices.Groups.Add(this.grpInvoices);
            this.LnInvoices.Groups.Add(this.grpClients);
            this.LnInvoices.Label = "FoxVoice";
            this.LnInvoices.Name = "LnInvoices";
            // 
            // grpInvoices
            // 
            this.grpInvoices.Items.Add(this.btnValidatePostalCode);
            this.grpInvoices.Items.Add(this.btn_FillClasificators);
            this.grpInvoices.Items.Add(this.btnLoadInvoiceList);
            this.grpInvoices.Items.Add(this.btnSaveInvoice);
            this.grpInvoices.Items.Add(this.btnNewFromThis);
            this.grpInvoices.Items.Add(this.btnOpenInvoice);
            this.grpInvoices.Items.Add(this.btnOpenTemplate);
            this.grpInvoices.Items.Add(this.btnGenerateInvoice);
            this.grpInvoices.Label = "Invoices";
            this.grpInvoices.Name = "grpInvoices";
            // 
            // btn_FillClasificators
            // 
            this.btn_FillClasificators.Label = "Fill Clasificators";
            this.btn_FillClasificators.Name = "btn_FillClasificators";
            this.btn_FillClasificators.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_FillClasificators_Click);
            // 
            // btnLoadInvoiceList
            // 
            this.btnLoadInvoiceList.Image = ((System.Drawing.Image)(resources.GetObject("btnLoadInvoiceList.Image")));
            this.btnLoadInvoiceList.Label = "Load Invoice List";
            this.btnLoadInvoiceList.Name = "btnLoadInvoiceList";
            this.btnLoadInvoiceList.ShowImage = true;
            this.btnLoadInvoiceList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadInvoiceList_Click);
            // 
            // btnSaveInvoice
            // 
            this.btnSaveInvoice.Label = "";
            this.btnSaveInvoice.Name = "btnSaveInvoice";
            // 
            // btnNewFromThis
            // 
            this.btnNewFromThis.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewFromThis.Image = ((System.Drawing.Image)(resources.GetObject("btnNewFromThis.Image")));
            this.btnNewFromThis.Label = "New From This";
            this.btnNewFromThis.Name = "btnNewFromThis";
            this.btnNewFromThis.ShowImage = true;
            this.btnNewFromThis.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewFromThis_Click);
            // 
            // btnOpenInvoice
            // 
            this.btnOpenInvoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenInvoice.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenInvoice.Image")));
            this.btnOpenInvoice.Label = "Open Invoice";
            this.btnOpenInvoice.Name = "btnOpenInvoice";
            this.btnOpenInvoice.ShowImage = true;
            this.btnOpenInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenInvoice_Click);
            // 
            // btnOpenTemplate
            // 
            this.btnOpenTemplate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenTemplate.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenTemplate.Image")));
            this.btnOpenTemplate.Label = "Open Template";
            this.btnOpenTemplate.Name = "btnOpenTemplate";
            this.btnOpenTemplate.ShowImage = true;
            this.btnOpenTemplate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenTemplate_Click);
            // 
            // btnGenerateInvoice
            // 
            this.btnGenerateInvoice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnGenerateInvoice.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerateInvoice.Image")));
            this.btnGenerateInvoice.Label = "Generate Invoice";
            this.btnGenerateInvoice.Name = "btnGenerateInvoice";
            this.btnGenerateInvoice.ShowImage = true;
            this.btnGenerateInvoice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnGenerateInvoice_Click);
            // 
            // grpClients
            // 
            this.grpClients.DialogLauncher = ribbonDialogLauncherImpl1;
            this.grpClients.Items.Add(this.btnNewClient);
            this.grpClients.Items.Add(this.btnLoadClientList);
            this.grpClients.Items.Add(this.btnUpdateClientInfo);
            this.grpClients.Label = "Clients";
            this.grpClients.Name = "grpClients";
            // 
            // btnNewClient
            // 
            this.btnNewClient.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNewClient.Image = ((System.Drawing.Image)(resources.GetObject("btnNewClient.Image")));
            this.btnNewClient.Label = "New Client";
            this.btnNewClient.Name = "btnNewClient";
            this.btnNewClient.ShowImage = true;
            this.btnNewClient.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewClient_Click);
            // 
            // btnLoadClientList
            // 
            this.btnLoadClientList.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnLoadClientList.Image = ((System.Drawing.Image)(resources.GetObject("btnLoadClientList.Image")));
            this.btnLoadClientList.Label = "Load Client List";
            this.btnLoadClientList.Name = "btnLoadClientList";
            this.btnLoadClientList.ShowImage = true;
            this.btnLoadClientList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadClientList_Click);
            // 
            // btnUpdateClientInfo
            // 
            this.btnUpdateClientInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateClientInfo.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateClientInfo.Image")));
            this.btnUpdateClientInfo.Label = "Update Client Info";
            this.btnUpdateClientInfo.Name = "btnUpdateClientInfo";
            this.btnUpdateClientInfo.ShowImage = true;
            this.btnUpdateClientInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateClientInfo_Click);
            // 
            // btnValidatePostalCode
            // 
            this.btnValidatePostalCode.Label = "Validate Postal Code";
            this.btnValidatePostalCode.Name = "btnValidatePostalCode";
            this.btnValidatePostalCode.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnValidatePostalCode_Click);
            // 
            // RibbonInvoice
            // 
            this.Name = "RibbonInvoice";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.LnInvoices);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonInvoice_Load);
            this.LnInvoices.ResumeLayout(false);
            this.LnInvoices.PerformLayout();
            this.grpInvoices.ResumeLayout(false);
            this.grpInvoices.PerformLayout();
            this.grpClients.ResumeLayout(false);
            this.grpClients.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab LnInvoices;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpInvoices;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewFromThis;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenTemplate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnGenerateInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpClients;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewClient;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadClientList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateClientInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenInvoice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadInvoiceList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_FillClasificators;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnValidatePostalCode;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonInvoice RibbonInvoice
        {
            get { return this.GetRibbon<RibbonInvoice>(); }
        }
    }
}
