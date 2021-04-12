using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmbeddedToAttachment {
    public static class Common {

        // Create private fields
        public static string appFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Night Owl Software\Outlook Add-ins\Embedded Image to Attachment Add-in");
        public static string imageCacheFolder = Path.Combine(appFolder, @"image-cache");     // Directory for storing temporary images while converting them to attachments
        public static string settingsFile = Path.Combine(appFolder, "settings.xml");         // A file used to store user-specified settings, for future implementation
        public static string blacklistFile = Path.Combine(appFolder, "blacklist.xml");       // A file used to store email addresses that this add-in will ignore, for future 
        private static Outlook.MailItem activeMailItem = null;
        private static bool buttonAvailable = false;

        /// <summary>
        /// Attaches the provided file to the provided Outlook MailItem as a regular attachment
        /// </summary>
        /// <param name="mail">OUtlook MailItem to attach files to</param>
        /// <param name="fileName">Full File Path of item to add to MailItem</param>
        public static void AddAttachment(Outlook.MailItem mail, string fileName) {
            if(fileName.Length > 0) {
                mail.Attachments.Add($"{imageCacheFolder}\\{fileName}", Outlook.OlAttachmentType.olByValue, mail.Attachments.Count + 1, fileName);
            }
        }
        public static void CheckForEmbeddedImages() {
            if(buttonAvailable == true) {
                // Check for any attachments (embedded or otherwise)
                if(activeMailItem.Attachments.Count > 0) {

                    // Create a blank list to store embedded image files into for iteration later
                    List<Outlook.Attachment> embeddedImages = new List<Outlook.Attachment>();

                    foreach(Outlook.Attachment attachment in activeMailItem.Attachments) {
                        string fileName = attachment.FileName;

                        // Look for any image-specific file extensions in the attachments
                        if(fileName.ToUpper().Contains(".JPG") ||
                            fileName.ToUpper().Contains(".PNG") ||
                            fileName.ToUpper().Contains(".TIF") ||
                            fileName.ToUpper().Contains(".TIFF") ||
                            fileName.ToUpper().Contains(".BMP")) {

                            // Check if attachment is included in the HTML Body, rather than in the Attachments DIV section
                            // This means it is definitely an embedded image, so we need to make a copy of it
                            if(activeMailItem.HTMLBody.Contains($"cid:{fileName}")) {

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
                            AddAttachment(activeMailItem, fileName);
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
                    buttonAvailable = false;
                }
            }
        }
        public static void CheckForEmbeddedImages(Outlook.MailItem mail) {
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

        public static void SetActiveMailItem(Outlook.MailItem mail) {
            activeMailItem = mail;
            buttonAvailable = true;
        }
        
    }
}
