using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmbeddedToAttachment {
    public partial class EmbeddedToAttachmentRibbonButton {
        private void EmbeddedToAttachmentRibbonButton_Load(object sender, RibbonUIEventArgs e) {

        }

        private void btnEmbedToAttach_Click(object sender, RibbonControlEventArgs e) {
            Common.CheckForEmbeddedImages();
        }
    }
}
