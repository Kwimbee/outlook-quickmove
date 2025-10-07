using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace QuickMove
{
    public partial class ThisAddIn
    {
        public static string version = "2.1.0";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        // call this to invoke the add-in
        public void quickMoveCalled()
        {
            Outlook.Folder current_folder = this.Application.ActiveExplorer().CurrentFolder as Outlook.Folder;

            // if is table view
            if (current_folder.CurrentView.ViewType == Outlook.OlViewType.olTableView)
            {
                Outlook.TableView current_view = current_folder.CurrentView as Outlook.TableView;
                if (current_view.ShowConversationByDate == true)
                {
                    Outlook.Folder rootFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;

                    List<string> allfolders = ListAllFolders(rootFolder);

                    var form = new FolderSelector(allfolders);
                    if (form.ShowDialog() == true)
                    {
                        string selectedFolder = form.selectedFolder;

                        Outlook.MAPIFolder foundFolder = findFolderByName(selectedFolder, rootFolder.Folders);

                        if (foundFolder != null)
                        {
                            Outlook.Selection selection = Application.ActiveExplorer().Selection;
                            Outlook.Selection convHeaders = selection.GetSelection(Outlook.OlSelectionContents.olConversationHeaders) as Outlook.Selection;

                            if (convHeaders.Count >= 1)
                            {
                                foreach (Outlook.ConversationHeader convHeader in convHeaders)
                                {
                                    Outlook.SimpleItems items = convHeader.GetItems();
                                    for (int i = 1; i <= items.Count; i++)
                                    {
                                        if (items[i] is Outlook.MailItem)
                                        {
                                            Outlook.MailItem mail = items[i] as Outlook.MailItem;

                                            try
                                            {
                                                mail.Move(foundFolder);
                                            }
                                            catch (System.Exception ex)
                                            {
                                                MessageBox.Show(ex.Message);
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                // Move items that are not conversation items
                                foreach (MailItem item in selection)
                                {
                                    try
                                    {
                                        item.Move(foundFolder);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        // call this to invoke quickbrowse
        public void quickBrowseCalled()
        {
            Outlook.Folder current_folder = this.Application.ActiveExplorer().CurrentFolder as Outlook.Folder;

            // if is table view
            if (current_folder.CurrentView.ViewType == Outlook.OlViewType.olTableView)
            {
                Outlook.TableView current_view = current_folder.CurrentView as Outlook.TableView;
                if (current_view.ShowConversationByDate == true)
                {
                    Outlook.Folder rootFolder = Application.Session.DefaultStore.GetRootFolder() as Outlook.Folder;

                    List<string> allfolders = ListAllFolders(rootFolder);

                    var form = new FolderSelector(allfolders);
                    if (form.ShowDialog() == true)
                    {
                        string selectedFolder = form.selectedFolder;

                        Outlook.MAPIFolder foundFolder = findFolderByName(selectedFolder, rootFolder.Folders);

                        if (foundFolder != null)
                        {
                            this.Application.ActiveExplorer().CurrentFolder = foundFolder;
                        }
                    }
                }
            }
        }

        // search folder by name
        private MAPIFolder findFolderByName(string folderPath, Folders folders)
        {
            string dir = folderPath.Substring(0, folderPath.Substring(4).IndexOf("\\") + 4);
            try
            {
                foreach (MAPIFolder mf in folders)
                {
                    if (!(mf.FullFolderPath.StartsWith(dir))) continue;
                    if (mf.FullFolderPath == folderPath) return mf;
                    else
                    {
                        MAPIFolder temp = findFolderByName(folderPath, mf.Folders);
                        if (temp != null)
                            return temp;
                    }
                }
                return null;
            }
            catch { return null; }
        }

        // Returns a list of strings of all folders in a mailbox folder
        public static List<string> ListAllFolders(Outlook.MAPIFolder rootFolder)
        {
            List<string> folderNames = new List<string>();

            // Call a recursive function to list all subfolders.
            ListSubfoldersNoIndent(rootFolder, folderNames);

            return folderNames;
        }

        private static void ListSubfoldersNoIndent(Outlook.MAPIFolder folder, List<string> folderNames)
        {
            // Add the name of the folder to the list.
            if (folder.DefaultItemType == OlItemType.olMailItem)
            {
                folderNames.Add(folder.FolderPath);
            }

            // Call this function recursively for each subfolder.
            foreach (Outlook.MAPIFolder subfolder in folder.Folders)
            {
                ListSubfoldersNoIndent(subfolder, folderNames);
            }
        }

        // creates the ribbon in the outlook ui
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new CustomRibbon();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Code généré par VSTO

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
