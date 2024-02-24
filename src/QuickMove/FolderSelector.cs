using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuickMove
{
    public partial class frmFolderSelector : Form
    {
        public string selectedFolder = "";
        
        private List<string> originalFolders = new List<string>();

        public frmFolderSelector(List<string> folders)
        {
            InitializeComponent();
            originalFolders = folders;
            lbFolders.Items.AddRange(originalFolders.ToArray());
        }

        // on load ui
        private void frmFolderSelector_Load(object sender, EventArgs e)
        {
            lbFolders.SelectedIndex = 0;
            txtSearch.Focus();

#if DEBUG
            this.Text = "QuickMove DEBUG - Folders";
                
            toolStripLabelVersion.Text = ThisAddIn.version + " - DEBUG";
#else

            toolStripLabelVersion.Text = ThisAddIn.version;
#endif

        }

        // On double click to item
        private void lbFolders_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            returnSelectedFolder();
        }

        // On text changed in search bar
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            lbFolders.Items.Clear();

            string searchText = txtSearch.Text.ToLower();

            foreach (string folder in originalFolders) {

                string folder_clean =  folder.ToLower().Replace('é', 'e').Replace('à', 'a').Replace('è','e').Replace('ê','e').Replace('ç', 'c').Replace('ù', 'u');
                string searchText_clean = searchText.ToLower().Replace('é', 'e').Replace('à', 'a').Replace('è', 'e').Replace('ê', 'e').Replace('ç', 'c').Replace('ù', 'u');

                if (folder_clean.Contains(searchText_clean))
                {
                    lbFolders.Items.Add(folder);
                }
            }

            if (lbFolders.Items.Count > 0)
            {
                lbFolders.SelectedIndex = 0;
            }
        }
        
        // handle key presses
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.Close();
                return true;
            }

            if (keyData == Keys.Down)
            {
                if (lbFolders.SelectedItems.Count == 0)
                {
                    lbFolders.SelectedIndex = 0;
                } else
                {
                    if (lbFolders.SelectedIndex == lbFolders.Items.Count - 1)
                    {
                        lbFolders.SelectedIndex = 0;
                    } else
                    {
                        lbFolders.SelectedIndex++;
                    }
                }
            }

            if (keyData == Keys.Up)
            {
                if (lbFolders.SelectedItems.Count == 0)
                {
                    lbFolders.SelectedIndex = 0;
                }
                else
                {
                    if (lbFolders.SelectedIndex == 0)
                    {
                        lbFolders.SelectedIndex = lbFolders.Items.Count - 1;
                    } else
                    {
                        lbFolders.SelectedIndex--;
                    }

                }
            }

            if (keyData == Keys.Enter)
            {
                returnSelectedFolder();
            }

            if (keyData == Keys.Home)
            {
                lbFolders.SelectedIndex = 0;
            }

            if (keyData == Keys.End)
            {
                lbFolders.SelectedIndex = lbFolders.Items.Count - 1;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        // returns the selected folder and close window
        private void returnSelectedFolder()
        {
            int index = this.lbFolders.SelectedIndex;
            if (index != System.Windows.Forms.ListBox.NoMatches)
            {
                this.DialogResult = DialogResult.OK;
                this.selectedFolder = lbFolders.Items[index].ToString();
                this.Close();
            }
        }
    }
}
