using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmbeddedToAttachment
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        Outlook.Explorer currentExplorer = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Create the AppData and ImageCache folders, if needed
            Directory.CreateDirectory(Common.appFolder);
            Directory.CreateDirectory(Common.imageCacheFolder);

            // Get a List of all files still in the Image Cache
            string[] allFiles = Directory.GetFiles(Common.imageCacheFolder);

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
                // EVENT IS CURRENTLY DISABLED, AUTOMATED CONVERSION CAUSING TOO MUCH MEMORY OVERHEAD
            //items = inbox.Items;
            //items.ItemAdd +=
            //    new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);

            // Set Event to Trigger any time a new window takes focus
            currentExplorer = Application.ActiveExplorer();
            currentExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (ExplorerChanged);
        }

        void Items_ItemAdd(object Item) {
            // Grab the received mail message
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            
            if(Item != null) {
                Common.CheckForEmbeddedImages(mail);
            }
        }

        void ExplorerChanged() {
            try {
                if(Application.ActiveExplorer().Selection.Count > 0) {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];

                    if(selObject is Outlook.MailItem) {
                        Outlook.MailItem mailItem =
                            (selObject as Outlook.MailItem);
                        Common.SetActiveMailItem(mailItem);
                    }
                }
            }
            catch(Exception e) {
                MessageBox.Show(e.Message);
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
