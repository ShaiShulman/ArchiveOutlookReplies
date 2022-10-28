using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace ArchiveOutlookReplies
{
    public partial class ThisAddIn
    {
        string sourceFolderID = null; // ID of folder from which replies will be copied. Copy the value that appear when you open a new explorer with that foldere. 
        //string targetFolderID = null; // ID the folder to which messages will be copies. If null, the local Drafts folder will be used. 

        Outlook.Explorers explorers;
        MAPIFolder targetFolder; // target folder to which messages will be copies. Default is the local Drafts folder.
        Dictionary<MailItem, Outlook.ItemEvents_10_SendEventHandler> sendEvents = 
            new Dictionary<MailItem, ItemEvents_10_SendEventHandler>(); // stores Send event handlers for unsubscribing 

        private void SelectFolders() {
            MessageBox.Show("Select source folder first following by target folder");
            sourceFolderID = Application.Session.PickFolder().EntryID;
            targetFolder = Application.Session.PickFolder();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SelectFolders();
            explorers = this.Application.Explorers;
            SetupNewExplorer(Application.ActiveExplorer()); // listen to active explorer
            explorers.NewExplorer += SetupNewExplorer; // listen to any new explorer created
            
        }

        // adds event listener for the reply event for each selected mail item if the explorer has the source folder as the active folder
        private void SetupNewExplorer(Explorer explorer)
        {
            explorer.SelectionChange += (() => {
                if (explorer.CurrentFolder.EntryID==this.sourceFolderID)
                {
                    foreach (var item in explorer.Selection)
                    {
                        MailItem mail = item as MailItem;
                        ((Outlook.ItemEvents_10_Event)mail).Reply += MailItem_Reply; // listen to reply event
                        ((Outlook.ItemEvents_10_Event)mail).ReplyAll += MailItem_Reply; // listen to reply-all event
                    }
                }
            });
        }

        // once user click Reply, listen to Send event (b/c we want to copy to email only once it is sent)
        private void MailItem_Reply(object Response, ref bool Cancel)
        {
            MailItem mail = Response as MailItem;
            Outlook.ItemEvents_10_SendEventHandler sendEventHandler = ((ref bool cancel) =>
            {
                ((Outlook.ItemEvents_10_Event)mail).Send -= sendEvents[mail];
                MailItem copy = mail.Copy(); // PROBLEM: Copy method doesn't work for inline editing :(
                // PROBLEM: currently the Reply event is fired multiple times for each sent (no idea why) and as a result mail message is copied multiple times. Need to find a way to track copied mail and unsubcribe from the event
                copy.Move(targetFolder);
                MessageBox.Show("Mail item '" + (Response as MailItem).Subject + "' copied to folder " + targetFolder.Name);
            });

            sendEvents[mail] = sendEventHandler;
            ((ItemEvents_10_Event)mail).Send += sendEventHandler;
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
