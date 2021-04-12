using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace EmbeddedToAttachment
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            

            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        void items_ItemAdd(object Item) {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if(Item != null) {
                
                if(mail.Attachments.Count > 0) {

                    if(Directory.Exists(@"C:\Temp") == false) {
                        Directory.CreateDirectory(@"C:\Temp");
                    }

                    List<Outlook.Attachment> imageAttachments = new List<Outlook.Attachment>();

                    foreach(Outlook.Attachment attachment in mail.Attachments) {

                        if(attachment.FileName.ToUpper().Contains(".JPG") ||
                            attachment.FileName.ToUpper().Contains(".PNG") ||
                            attachment.FileName.ToUpper().Contains(".TIF") ||
                            attachment.FileName.ToUpper().Contains(".TIFF") ||
                            attachment.FileName.ToUpper().Contains(".BMP")) {

                            imageAttachments.Add(attachment);
                        }
                    }

                    foreach(Outlook.Attachment image in imageAttachments) {
                        string fileName = image.FileName;
                        image.SaveAsFile($"C:\\Temp\\{fileName}");
                        AddAttachment(mail, fileName);
                    }

                    foreach(Outlook.Attachment image in imageAttachments) {
                        string fileName = image.FileName;
                        File.Delete($"C:\\Temp\\{fileName}");
                    }
                }
            }

        }

        private void AddAttachment(Outlook.MailItem mail, string fileName) {
            if(fileName.Length > 0) {
                mail.Attachments.Add($"C:\\Temp\\{fileName}", Outlook.OlAttachmentType.olByValue, mail.Attachments.Count + 1, fileName);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
