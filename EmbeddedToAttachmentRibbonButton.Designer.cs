
namespace EmbeddedToAttachment {
    partial class EmbeddedToAttachmentRibbonButton : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public EmbeddedToAttachmentRibbonButton()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.tabEmbedAttach = this.Factory.CreateRibbonTab();
            this.grpEmbedToAttach = this.Factory.CreateRibbonGroup();
            this.btnEmbedToAttach = this.Factory.CreateRibbonButton();
            this.tabEmbedAttach.SuspendLayout();
            this.grpEmbedToAttach.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEmbedAttach
            // 
            this.tabEmbedAttach.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEmbedAttach.ControlId.OfficeId = "TabReadMessage";
            this.tabEmbedAttach.Groups.Add(this.grpEmbedToAttach);
            this.tabEmbedAttach.Label = "TabReadMessage";
            this.tabEmbedAttach.Name = "tabEmbedAttach";
            // 
            // grpEmbedToAttach
            // 
            this.grpEmbedToAttach.Items.Add(this.btnEmbedToAttach);
            this.grpEmbedToAttach.Label = "Embed To Attach";
            this.grpEmbedToAttach.Name = "grpEmbedToAttach";
            // 
            // btnEmbedToAttach
            // 
            this.btnEmbedToAttach.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEmbedToAttach.Image = global::EmbeddedToAttachment.Properties.Resources.EmbedToAttach;
            this.btnEmbedToAttach.Label = "Embed to Attach";
            this.btnEmbedToAttach.Name = "btnEmbedToAttach";
            this.btnEmbedToAttach.ShowImage = true;
            this.btnEmbedToAttach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmbedToAttach_Click);
            // 
            // EmbeddedToAttachmentRibbonButton
            // 
            this.Name = "EmbeddedToAttachmentRibbonButton";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tabEmbedAttach);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.EmbeddedToAttachmentRibbonButton_Load);
            this.tabEmbedAttach.ResumeLayout(false);
            this.tabEmbedAttach.PerformLayout();
            this.grpEmbedToAttach.ResumeLayout(false);
            this.grpEmbedToAttach.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEmbedAttach;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEmbedToAttach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEmbedToAttach;
    }

    partial class ThisRibbonCollection {
        internal EmbeddedToAttachmentRibbonButton EmbeddedToAttachmentRibbonButton {
            get { return this.GetRibbon<EmbeddedToAttachmentRibbonButton>(); }
        }
    }
}
