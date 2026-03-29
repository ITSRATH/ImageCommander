using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace ImageCommander
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    /// 
    using WpfMsg = System.Windows.MessageBox;


    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        public class FileItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set { _isSelected = value; OnPropertyChanged(); }
            }

            public string FileName { get; set; } = "";
            public string FileType { get; set; } = "";
            public string FileSize { get; set; } = "";

            public event PropertyChangedEventHandler? PropertyChanged;
            protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private static string FormatSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len /= 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        private void BrowseSourceFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog src = new FolderBrowserDialog();
            src.ShowDialog();
            sourceFolderPath.Text = src.SelectedPath;
            LoadFilesFromPath(src.SelectedPath);
        }
        private void LoadFilesFromPath(string folderPath)
        {
            var items = new List<FileItem>();

            if (Directory.Exists(folderPath))
            {
                var extensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp" };
                foreach (var file in Directory.EnumerateFiles(folderPath))
                {
                    var ext = System.IO.Path.GetExtension(file).ToLowerInvariant();
                    if (extensions.Contains(ext))
                    {
                        var fi = new FileInfo(file);
                        items.Add(new FileItem
                        {
                            FileName = fi.Name,
                            FileType = ext,
                            FileSize = FormatSize(fi.Length)
                        });
                    }
                }
            }

            // Binden an ListView
            lvFiles.ItemsSource = items;
        }
        private void BrowseDestinationFolder_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog dst = new FolderBrowserDialog();
            dst.ShowDialog();
            destinationFolderPath.Text = dst.SelectedPath;
        }

        private void startButton_Click(object sender, RoutedEventArgs e)
        {
            var srcpath = sourceFolderPath.Text;
            var destpath = destinationFolderPath.Text;
            var ratio = TypeConvertionTo.SelectionBoxItem;
            var quality = qualitySelection.SelectionBoxItem;
            var rename = renameActive.IsChecked;
            var renameTxt = renameText.Text;
            var removeMeta = removeMetadata.IsChecked;
            WpfMsg.Show("Convertion Ratio: " + ratio + " Quality: " + quality + " Remove Meta: " + removeMeta + " Src: " + srcpath + " Dest: " + destpath);
            if (rename == true)
            {
                WpfMsg.Show("Rename is active, rename text: " + renameTxt);
            }
        }

    }
}