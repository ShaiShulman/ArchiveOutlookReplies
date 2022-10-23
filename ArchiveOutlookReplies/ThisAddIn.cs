using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using static System.Collections.Specialized.BitVector32;

namespace ArchiveOutlookReplies
{
    public partial class ThisAddIn
    {
        readonly string sourceFolderID = "000000005DD9BCB153A2E847B2EA167045ADB2A50100B84A13FF84FE484EA1A67250243D94D80029AEA7113B0000"; // ID of folder from which replies will be copied. Copy the value that appear when you open a new explorer with that foldere. 
        readonly string targetFolderID = null; // ID the folder to which messages will be copies. If null, the local Drafts folder will be used. 

        Outlook.Explorers explorers;
        MAPIFolder targetFolder; // target folder to which messages will be copies. Default is the local Drafts folder.
        Dictionary<MailItem, Outlook.ItemEvents_10_SendEventHandler> sendEvents = 
            new Dictionary<MailItem, ItemEvents_10_SendEventHandler>(); // stores Send event handlers for unsubscribing 

        ProcessedMailItems List<MailItem> = List<MailItem>();

        // get MAPIFolder object for target folder (from ID if specified, otherwise get local Drafts folder)
        private MAPIFolder getTargetFolder() {
            Microsoft.Office.Interop.Outlook.NameSpace session = Application.Session;
            if (targetFolderID == null)
            {
                return session.GetDefaultFolder(OlDefaultFolders.olFolderDrafts);
            }
            else
                return session.GetFolderFromID(targetFolderID);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            targetFolder = getTargetFolder();
            explorers = this.Application.Explorers;
            setupNewExplorer(Application.ActiveExplorer()); // listen to active explorer
            explorers.NewExplorer += setupNewExplorer; // listen to any new explorer created
            
        }

        // adds event listener for the reply event for each selected mail item if the explorer has the source folder as the active folder
        private void setupNewExplorer(Explorer explorer)
        {
            MessageBox.Show("New explorer with FolderID:" + explorer.CurrentFolder.EntryID);
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
