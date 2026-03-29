using System;
using System.Windows;
using WinForms = System.Windows.Forms;

namespace ImageCommander
{
    public partial class SettingsWindow : Window
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        // controls are available by their x:Name fields generated from XAML

        private void BrowseWatermark_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new WinForms.OpenFileDialog())
            {
                dlg.Filter = "PNG Files|*.png|All Files|*.*";
                dlg.Multiselect = false;
                var res = dlg.ShowDialog();
                if (res == WinForms.DialogResult.OK)
                {
                    tbWatermarkFile.Text = dlg.FileName;
                }
            }
        }

        private void BrowseSource_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "Quellordner wählen";
                dlg.ShowNewFolderButton = true;
                var res = dlg.ShowDialog();
                if (res == WinForms.DialogResult.OK)
                {
                    tbDefaultSource.Text = dlg.SelectedPath;
                }
            }
        }

        private void BrowseDestination_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "Zielordner wählen";
                dlg.ShowNewFolderButton = true;
                var res = dlg.ShowDialog();
                if (res == WinForms.DialogResult.OK)
                {
                    tbDefaultDestination.Text = dlg.SelectedPath;
                }
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            // Optionally validate values here
            this.DialogResult = true;
            this.Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
