using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace QuickMove
{
    /// <summary>
    /// Interaction logic for FolderSelector.xaml
    /// </summary>
    public partial class FolderSelector : Window
    {
        public string selectedFolder = "";

        private string mailBoxName = "";

        private List<string> originalFolders = new List<string>();

        public FolderSelector(List<string> folders)
        {
            InitializeComponent();


            if (IsWindowsInDarkMode())
            {
                ApplyDarkTheme();
            }

            originalFolders = folders;

            mailBoxName = originalFolders[0] + "\\";

            for (int i = 0; i < originalFolders.Count; i++)
            {
                string folder_name = originalFolders[i].Replace(mailBoxName, "");
                lbFolders.Items.Add(folder_name);
            }

            lbFolders.SelectedIndex = 0;
            txtSearch.Focus();
#if DEBUG
            this.Title = "QuickMove " + ThisAddIn.version + " - DEBUG";
#else
            this.Title = "QuickMove " + ThisAddIn.version;
#endif
        }

        private void txtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            lbFolders.Items.Clear();

            string searchText = txtSearch.Text.ToLower();

            foreach (string folder in originalFolders)
            {

                string folder_clean = folder.ToLower().Replace('é', 'e').Replace('à', 'a').Replace('è', 'e').Replace('ê', 'e').Replace('ç', 'c').Replace('ù', 'u');
                string searchText_clean = searchText.ToLower().Replace('é', 'e').Replace('à', 'a').Replace('è', 'e').Replace('ê', 'e').Replace('ç', 'c').Replace('ù', 'u');

                if (folder_clean.Contains(searchText_clean))
                {
                    lbFolders.Items.Add(folder.Replace(mailBoxName, ""));
                }
            }

            if (lbFolders.Items.Count > 0)
            {
                lbFolders.SelectedIndex = 0;
            }
        }

        private void lbFolders_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            returnSelectedFolder();
        }

        // returns the selected folder and close window
        private void returnSelectedFolder()
        {
            int index = this.lbFolders.SelectedIndex;
            this.DialogResult = true;
            this.selectedFolder = mailBoxName + lbFolders.Items[index].ToString();
            this.Close();
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
            }

            if (e.Key == Key.Down)
            {
                if (lbFolders.SelectedItems.Count == 0)
                {
                    lbFolders.SelectedIndex = 0;
                }
                else
                {
                    if (lbFolders.SelectedIndex == lbFolders.Items.Count - 1)
                    {
                        lbFolders.SelectedIndex = 0;
                    }
                    else
                    {
                        lbFolders.SelectedIndex++;
                    }
                }
            }

            if (e.Key == Key.Up)
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
                    }
                    else
                    {
                        lbFolders.SelectedIndex--;
                    }

                }
            }

            if (e.Key == Key.Enter)
            {
                returnSelectedFolder();
            }

            if (e.Key == Key.Home)
            {
                lbFolders.SelectedIndex = 0;
            }

            if (e.Key == Key.End)
            {
                lbFolders.SelectedIndex = lbFolders.Items.Count - 1;
            }
        }

        private void lbFolders_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            lbFolders.ScrollIntoView(lbFolders.SelectedItem);
        }

        private void ApplyDarkTheme()
        {
            SolidColorBrush DarkDark = new SolidColorBrush(Color.FromRgb(31, 31, 31));
            SolidColorBrush Dark = new SolidColorBrush(Color.FromRgb(41, 41, 41));
            SolidColorBrush DarkModeText = new SolidColorBrush(Color.FromRgb(255,255,255));

            lbFolders.Background = DarkDark;
            lbFolders.Foreground = DarkModeText;
            
            txtSearch.Background = DarkDark;
            txtSearch.Foreground = DarkModeText;
        }

        private bool IsWindowsInDarkMode()
        {
            const string registryKey = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string registryValue = "AppsUseLightTheme";

            object theme = Registry.GetValue(registryKey, registryValue, 1);
            return theme is int value && value == 0; // 0 = Dark Mode, 1 = Light Mode
        }
    }
}
