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
        string appFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Night Owl Software\Outlook Add-ins\Embedded Image to Attachment Add-in");
        string imageCacheFolder;
        string settingsFile;
        string blacklistFile;

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create our base directory and file paths
            imageCacheFolder = Path.Combine(appFolder, @"image-cache");     // Directory for storing temporary images while converting them to attachments
            settingsFile = Path.Combine(appFolder, "settings.xml");         // A file used to store user-specified settings, for future implementation
            blacklistFile = Path.Combine(appFolder, "blacklist.xml");       // A file used to store email addresses that this add-in will ignore, for future implementation

            // Create the AppData and ImageCache folders, if needed
            Directory.CreateDirectory(appFolder);
            Directory.CreateDirectory(imageCacheFolder);

            // Get a List of all files still in the Image Cache
            string[] allFiles = Directory.GetFiles(imageCacheFolder);

            // Clear the Image Cache, if any images exist
            if(allFiles.Length > 0) {
                for(int i = allFiles.Length - 1; i >= 0; i--) {
                    File.Delete(allFiles[i]);
                }
            }

            // Get the Outlook Namespace so we can monitor incoming messages
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Outlook.OlDefaultFolders.olFolderInbox);

            // Get a list of inbox items and add an event flag for when new items are added (a message is received)
            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        void items_ItemAdd(object Item) {
            Outlook.MailItem mail = (Outlook.MailItem)Item;

            if(Item != null) {                
                if(mail.Attachments.Count > 0) {

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

                    if(imageAttachments.Count > 0) {
                        string tempSubject = mail.Subject;

                        mail.Subject = $"[ATTACHMENTS] {tempSubject}";

                        foreach(Outlook.Attachment image in imageAttachments) {
                            string fileName = image.FileName;
                            image.SaveAsFile($"{imageCacheFolder}\\{fileName}");
                            AddAttachment(mail, fileName);
                        }

                        foreach(Outlook.Attachment image in imageAttachments) {
                            string fileName = image.FileName;
                            File.Delete($"{imageCacheFolder}\\{fileName}");
                        }
                    }
                }
            }

        }

        private void AddAttachment(Outlook.MailItem mail, string fileName) {
            if(fileName.Length > 0) {
                mail.Attachments.Add($"{imageCacheFolder}\\{fileName}", Outlook.OlAttachmentType.olByValue, mail.Attachments.Count + 1, fileName);
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
        }
        
        #endregion
    }
}
