using System;
using System.IO;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmbeddedToAttachment
{
    public partial class ThisAddIn
    {
        // Create private fields
        private string appFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Night Owl Software\Outlook Add-ins\Embedded Image to Attachment Add-in");
        private string imageCacheFolder;
        private string settingsFile;
        private string blacklistFile;

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
            // Grab the received mail message
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            
            if(Item != null) {

                CheckForEmbeddedImages(mail);

            }

        }

        /// <summary>
        /// Attaches the provided file to the provided Outlook MailItem as a regular attachment
        /// </summary>
        /// <param name="mail">OUtlook MailItem to attach files to</param>
        /// <param name="fileName">Full File Path of item to add to MailItem</param>
        private void AddAttachment(Outlook.MailItem mail, string fileName) {
            if(fileName.Length > 0) {
                mail.Attachments.Add($"{imageCacheFolder}\\{fileName}", Outlook.OlAttachmentType.olByValue, mail.Attachments.Count + 1, fileName);
            }
        }

        private void CheckForEmbeddedImages(Outlook.MailItem mail) {
            // Check for any attachments (embedded or otherwise)
            if(mail.Attachments.Count > 0) {

                // Create a blank list to store embedded image files into for iteration later
                List<Outlook.Attachment> embeddedImages = new List<Outlook.Attachment>();

                foreach(Outlook.Attachment attachment in mail.Attachments) {
                    string fileName = attachment.FileName;

                    // Look for any image-specific file extensions in the attachments
                    if(fileName.ToUpper().Contains(".JPG") ||
                        fileName.ToUpper().Contains(".PNG") ||
                        fileName.ToUpper().Contains(".TIF") ||
                        fileName.ToUpper().Contains(".TIFF") ||
                        fileName.ToUpper().Contains(".BMP")) {

                        // Check if attachment is included in the HTML Body, rather than in the Attachments DIV section
                        // This means it is definitely an embedded image, so we need to make a copy of it
                        if(mail.HTMLBody.Contains($"cid:{fileName}")) {

                            embeddedImages.Add(attachment);

                        }
                    }
                }

                // If we found embedded images, lets iterate through them
                if(embeddedImages.Count > 0) {

                    // First, iterate through and save all the images to a local, temporary image-cache
                    // Then, using that cache, re-attach the images to the mail as regular attachments
                    foreach(Outlook.Attachment image in embeddedImages) {
                        string fileName = image.FileName;
                        image.SaveAsFile($"{imageCacheFolder}\\{fileName}");
                        AddAttachment(mail, fileName);
                    }

                    // Once the attachment process is finished, work throught he list again
                    // and delete all the image files we created
                    foreach(Outlook.Attachment image in embeddedImages) {
                        string fileName = image.FileName;
                        File.Delete($"{imageCacheFolder}\\{fileName}");
                    }
                }

                // Empty out our list for good measure, now that we're done
                embeddedImages.Clear();
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
