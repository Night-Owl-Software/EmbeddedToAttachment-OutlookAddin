
namespace EmbeddedToAttachment {
    partial class EmbeddedToAttachmentMessageRibbonButton : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public EmbeddedToAttachmentMessageRibbonButton()
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
            this.tabEmbedAttachMessage = this.Factory.CreateRibbonTab();
            this.grpEmbedToAttach = this.Factory.CreateRibbonGroup();
            this.btnEmbedToAttach = this.Factory.CreateRibbonButton();
            this.tabEmbedAttachMessage.SuspendLayout();
            this.grpEmbedToAttach.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabEmbedAttachMessage
            // 
            this.tabEmbedAttachMessage.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabEmbedAttachMessage.ControlId.OfficeId = "TabReadMessage";
            this.tabEmbedAttachMessage.Groups.Add(this.grpEmbedToAttach);
            this.tabEmbedAttachMessage.Label = "TabReadMessage";
            this.tabEmbedAttachMessage.Name = "tabEmbedAttachMessage";
            // 
            // grpEmbedToAttach
            // 
            this.grpEmbedToAttach.Items.Add(this.btnEmbedToAttach);
            this.grpEmbedToAttach.Label = "Embedded to Attached";
            this.grpEmbedToAttach.Name = "grpEmbedToAttach";
            // 
            // btnEmbedToAttach
            // 
            this.btnEmbedToAttach.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEmbedToAttach.Image = global::EmbeddedToAttachment.Properties.Resources.EmbedToAttach;
            this.btnEmbedToAttach.Label = "Convert...";
            this.btnEmbedToAttach.Name = "btnEmbedToAttach";
            this.btnEmbedToAttach.ShowImage = true;
            this.btnEmbedToAttach.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEmbedToAttach_Click);
            // 
            // EmbeddedToAttachmentMessageRibbonButton
            // 
            this.Name = "EmbeddedToAttachmentMessageRibbonButton";
            this.RibbonType = "Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tabEmbedAttachMessage);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.EmbeddedToAttachmentRibbonButton_Load);
            this.tabEmbedAttachMessage.ResumeLayout(false);
            this.tabEmbedAttachMessage.PerformLayout();
            this.grpEmbedToAttach.ResumeLayout(false);
            this.grpEmbedToAttach.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEmbedAttachMessage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpEmbedToAttach;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEmbedToAttach;
    }

    partial class ThisRibbonCollection {
        internal EmbeddedToAttachmentMessageRibbonButton EmbeddedToAttachmentRibbonButton {
            get { return this.GetRibbon<EmbeddedToAttachmentMessageRibbonButton>(); }
        }
    }
}
