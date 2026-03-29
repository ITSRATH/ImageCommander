using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using WinForms = System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Drawing.Imaging;


namespace ImageCommander
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    /// 
    using WpfMsg = System.Windows.MessageBox;


    public partial class MainWindow : Window
    {
        private AppSettings _settings;

        private static readonly string[] SupportedExtensions = new[] {
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp",
            ".heic", ".heif", ".raw", ".cr2", ".nef", ".arw", ".orf", ".dng", ".raf"
        };

        private static readonly string[] SizeUnits = { "B", "KB", "MB", "GB" };

        private string GetAppDataPath()
        {
            var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ImageCommander");
            Directory.CreateDirectory(folder);
            return folder;
        }

        private void MenuSettings_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SettingsWindow();
            // populate settings window fields
            try
            {
                dlg.tbDefaultSource.Text = _settings.DefaultSource ?? string.Empty;
                dlg.tbDefaultDestination.Text = _settings.DefaultDestination ?? string.Empty;
                dlg.chkWatermarkEnabled.IsChecked = _settings.WatermarkEnabled;
                dlg.tbWatermarkFile.Text = _settings.WatermarkFile ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("MenuSettings_Click populate error: " + ex);
            }

            var res = dlg.ShowDialog();
            if (res == true)
            {
                // read back
                _settings.DefaultSource = dlg.tbDefaultSource.Text ?? "";
                _settings.DefaultDestination = dlg.tbDefaultDestination.Text ?? "";
                _settings.WatermarkEnabled = dlg.chkWatermarkEnabled.IsChecked == true;
                _settings.WatermarkFile = dlg.tbWatermarkFile.Text ?? "";
                SaveSettings();

                // apply to UI
                if (!string.IsNullOrWhiteSpace(_settings.DefaultSource))
                {
                    sourceFolderPath.Text = _settings.DefaultSource;
                    // load files from the stored default source immediately
                try { LoadFilesFromPath(_settings.DefaultSource); } catch (Exception ex) { Debug.WriteLine("LoadFilesFromPath error: " + ex); }
                }
                if (!string.IsNullOrWhiteSpace(_settings.DefaultDestination)) destinationFolderPath.Text = _settings.DefaultDestination;
            }
        }

        private void MenuExit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private string SettingsFilePath => System.IO.Path.Combine(GetAppDataPath(), "settings.json");

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var json = File.ReadAllText(SettingsFilePath);
                    _settings = System.Text.Json.JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
                }
                else _settings = new AppSettings();
            }
            catch (Exception ex) { Debug.WriteLine("LoadSettings failed: " + ex); _settings = new AppSettings(); }

            // apply to UI
            if (!string.IsNullOrWhiteSpace(_settings.DefaultSource)) sourceFolderPath.Text = _settings.DefaultSource;
            if (!string.IsNullOrWhiteSpace(_settings.DefaultDestination)) destinationFolderPath.Text = _settings.DefaultDestination;
        }

        private void SaveSettings()
        {
            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(_settings, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SettingsFilePath, json);
            }
            catch (Exception ex) { Debug.WriteLine("SaveSettings failed: " + ex); }
        }

        public MainWindow()
        {
            InitializeComponent();
            LoadSettings();
            // If a default source folder is stored in settings, load its files immediately
            try
            {
                if (!string.IsNullOrWhiteSpace(_settings?.DefaultSource) && Directory.Exists(_settings.DefaultSource))
                {
                    sourceFolderPath.Text = _settings.DefaultSource;
                    LoadFilesFromPath(_settings.DefaultSource);
                }
            }
            catch (Exception ex) { Debug.WriteLine("Constructor LoadFilesFromPath error: " + ex); }
        }

        // Flexible loader that will try to use ImageMagick (Magick.NET) via reflection for
        // formats not natively supported by WPF. If Magick.NET is not available, falls
        // back to WIC BitmapDecoder. Returns null if image could not be loaded.
        private BitmapSource? LoadBitmapSourceFlexible(string path)
        {
            // Try ImageMagick (Magick.NET) via reflection so the project does not hard-depend on the package.
            try
            {
                // Try Q16 then Q8 assembly names
                var magickType = Type.GetType("ImageMagick.MagickImage, Magick.NET-Q16-AnyCPU")
                                ?? Type.GetType("ImageMagick.MagickImage, Magick.NET-Q8-AnyCPU");
                if (magickType != null)
                {
                    // Create MagickImage instance from file path
                    var img = Activator.CreateInstance(magickType, new object[] { path });
                    if (img != null)
                    {
                        try
                        {
                            using var ms = new MemoryStream();
                            // Prefer writing to PNG so WPF can decode it reliably
                            var writeMethod = magickType.GetMethod("Write", new Type[] { typeof(Stream), typeof(string) })
                                              ?? magickType.GetMethod("Write", new Type[] { typeof(Stream) });
                            if (writeMethod != null)
                            {
                                if (writeMethod.GetParameters().Length == 2)
                                    writeMethod.Invoke(img, new object[] { ms, "PNG" });
                                else
                                    writeMethod.Invoke(img, new object[] { ms });

                                ms.Position = 0;
                                var decoder = PngBitmapDecoder.Create(ms, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                                var frame = decoder.Frames[0];
                                if (frame.CanFreeze) frame.Freeze();
                                return frame;
                            }
                        }
                        finally
                        {
                            (img as IDisposable)?.Dispose();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("LoadBitmapSourceFlexible Magick attempt failed: " + ex);
            }

            // Fallback to WIC BitmapDecoder
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    BitmapDecoder decoder = BitmapDecoder.Create(fs, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                    var frame = decoder.Frames[0];
                    if (frame.CanFreeze) frame.Freeze();
                    return frame;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("LoadBitmapSourceFlexible WIC fallback failed: " + ex);
                return null;
            }
        }
        public class FileItem : INotifyPropertyChanged
        {
            private bool _isSelected;
            public bool IsSelected
            {
                get => _isSelected;
                set { _isSelected = value; OnPropertyChanged(); }
            }

            public string FileName { get; set; } = string.Empty;
            public string FullPath { get; set; } = string.Empty;
            public string FileType { get; set; } = string.Empty;
            public string FileDate { get; set; } = string.Empty;
            public string FileSize { get; set; } = string.Empty;

            public event PropertyChangedEventHandler? PropertyChanged;
            protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        private static string FormatSize(long bytes)
        {
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < SizeUnits.Length - 1)
            {
                order++;
                len /= 1024;
            }
            return $"{len:0.##} {SizeUnits[order]}";
        }

        private void BrowseSourceFolder_Click(object sender, RoutedEventArgs e)
        {
            using var src = new WinForms.FolderBrowserDialog();
            var res = src.ShowDialog();
            if (res == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(src.SelectedPath))
            {
                sourceFolderPath.Text = src.SelectedPath;
                try { LoadFilesFromPath(src.SelectedPath); } catch (Exception ex) { Debug.WriteLine("LoadFilesFromPath error: " + ex); }
            }
        }
        private void LoadFilesFromPath(string folderPath)
        {
            var items = new List<FileItem>();

            if (Directory.Exists(folderPath))
            {
                foreach (var file in Directory.EnumerateFiles(folderPath))
                {
                    var ext = Path.GetExtension(file).ToLowerInvariant();
                    if (SupportedExtensions.Contains(ext))
                    {
                        var fi = new FileInfo(file);
                        items.Add(new FileItem
                        {
                            FileName = fi.Name,
                            FullPath = fi.FullName,
                            FileType = ext.TrimStart('.').ToUpperInvariant() + "-Datei",
                            FileDate = fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm"),
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
            using var dst = new WinForms.FolderBrowserDialog();
            var res = dst.ShowDialog();
            if (res == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dst.SelectedPath))
                destinationFolderPath.Text = dst.SelectedPath;
        }

        private void selectAllCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (lvFiles.ItemsSource is System.Collections.IEnumerable items)
            {
                foreach (var obj in items)
                {
                    if (obj is FileItem fi)
                        fi.IsSelected = true;
                }
            }
        }

        private void selectAllCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            if (lvFiles.ItemsSource is System.Collections.IEnumerable items)
            {
                foreach (var obj in items)
                {
                    if (obj is FileItem fi)
                        fi.IsSelected = false;
                }
            }
        }

        private async void startButton_Click(object sender, RoutedEventArgs e)
        {
            var srcpath = sourceFolderPath.Text;
            var destpath = destinationFolderPath.Text;
            if (string.IsNullOrWhiteSpace(srcpath) || !Directory.Exists(srcpath))
            {
                WpfMsg.Show("Bitte gültigen Quellordner wählen.");
                return;
            }

            // Wenn kein Zielordner gesetzt ist oder er nicht existiert,
            // automatisch einen Unterordner im Quellordner erstellen.
            if (string.IsNullOrWhiteSpace(destpath) || !Directory.Exists(destpath))
            {
                try
                {
                    destpath = System.IO.Path.Combine(srcpath, "Konvertierungen - ImageCommander");
                    if (!Directory.Exists(destpath))
                        Directory.CreateDirectory(destpath);

                    // UI aktualisieren
                    Dispatcher.Invoke(() => destinationFolderPath.Text = destpath);
                }
                catch (Exception ex)
                {
                    WpfMsg.Show("Zielordner konnte nicht erstellt werden: " + ex.Message);
                    return;
                }
            }

            var targetItem = TypeConvertionTo.SelectedItem as ComboBoxItem;
            var targetFormat = (targetItem?.Content ?? "JPG").ToString();
            var qualityItem = qualitySelection.SelectedItem as ComboBoxItem;
            var qualityStr = (qualityItem?.Content ?? "100%").ToString();
            int quality = 100;
            if (qualityStr.Contains("%"))
            {
                var digits = new string(qualityStr.Where(char.IsDigit).ToArray());
                if (int.TryParse(digits, out var q)) quality = q;
            }

            var rename = renameActive.IsChecked == true;
            var renameTxt = renameText.Text;
            var removeMeta = removeMetadata.IsChecked == true;

            // Mail-Accept settings (read from UI before starting background task)
            var mailAcceptItem = mailAcceptSizeSelection.SelectedItem as ComboBoxItem;
            var mailAcceptStr = (mailAcceptItem?.Content ?? "Nein").ToString();
            // Consider both "Nein" and "Keine" as disabled. Enable only when a size (digits) exists.
            var digitsInMail = new string(mailAcceptStr.Where(char.IsDigit).ToArray());
            bool mailAccept = !string.IsNullOrWhiteSpace(digitsInMail);
            long mailThresholdBytes = 10L * 1024 * 1024; // default 10 MB
            if (long.TryParse(digitsInMail, out var mb))
            {
                mailThresholdBytes = mb * 1024L * 1024L;
            }

            var items = (lvFiles.ItemsSource as IEnumerable<FileItem>)?.Where(i => i.IsSelected).ToList();
            if (items == null || items.Count == 0)
            {
                WpfMsg.Show("Bitte mindestens eine Datei auswählen.");
                return;
            }

            // ZIP option
            var zipAfter = zipAfterConvert.IsChecked == true;
            var createdFiles = new List<string>();
            var createdFilesData = new List<(string name, byte[] data)>();

            // Disable UI
            startButton.IsEnabled = false;
            selectAllCheckBox.IsEnabled = false;
            statusProgressBar.Foreground = new SolidColorBrush(Colors.Blue);
            statusText.Text = "Starte Konvertierung...";

            await Task.Run(() =>
            {
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    try
                    {
                        Dispatcher.Invoke(() => statusText.Text = $"Verarbeite {i + 1}/{items.Count}: {item.FileName}");

                        // Load image using flexible loader (supports Magick.NET if available)
                        var loaded = LoadBitmapSourceFlexible(item.FullPath);
                        if (loaded == null)
                            throw new Exception("Konnte die Bilddatei nicht laden.");

                        BitmapSource frameToSave = loaded;
                        if (removeMeta)
                        {
                            // create a writeable copy to strip metadata
                            frameToSave = new WriteableBitmap(loaded);
                        }

                            // Only apply resizing/compression for files larger than the selected mail threshold.
                            FileInfo originalFi = new FileInfo(item.FullPath);
                            bool shouldCompressForMail = mailAccept && quality < 100 && originalFi.Length > mailThresholdBytes;

                            // Apply resizing based on selected percentage to approximate target file size.
                            double sizeRatio = Math.Max(1, Math.Min(100, quality)) / 100.0; // 1.0 means original
                            double scale = Math.Sqrt(sizeRatio);
                            BitmapSource finalSource = frameToSave;
                            if (shouldCompressForMail && scale > 0 && scale < 1.0)
                            {
                                var transform = new ScaleTransform(scale, scale);
                                var transformed = new TransformedBitmap(frameToSave, transform);
                                transformed.Freeze();
                                finalSource = transformed;
                            }

                            // We'll always encode to a memory buffer first so we can check the final size
                            string ext = ".jpg";
                            byte[] encodedData = null;

                            // Initial parameters for iterative compression if Mail-Accept requires it
                            int currentQuality = Math.Max(1, Math.Min(100, quality));
                            double currentScale = Math.Max(0.01, scale);

                            // Try encoding and, if Mail-Accept is enabled and result is larger than threshold,
                            // iteratively reduce quality and then scale until it fits under the threshold.
                            for (int attempt = 0; attempt < 12; attempt++)
                            {
                                BitmapSource sourceToEncode = finalSource;
                                if (currentScale > 0 && currentScale < 0.999)
                                {
                                    var transform = new ScaleTransform(currentScale, currentScale);
                                    var transformed = new TransformedBitmap(frameToSave, transform);
                                    transformed.Freeze();
                                    sourceToEncode = transformed;
                                }

                                BitmapEncoder encoderToUse;
                                var tf = targetFormat.ToUpperInvariant();
                                if (tf == "PNG")
                                {
                                    encoderToUse = new PngBitmapEncoder();
                                    ext = ".png";
                                }
                                else if (tf == "WEBP")
                                {
                                    Dispatcher.Invoke(() => WpfMsg.Show("WebP wird nicht nativ unterstützt. Es wird PNG als Ersatz verwendet."));
                                    encoderToUse = new PngBitmapEncoder();
                                    ext = ".png";
                                }
                                else if (tf == "BMP")
                                {
                                    encoderToUse = new BmpBitmapEncoder();
                                    ext = ".bmp";
                                }
                                else if (tf == "TIFF")
                                {
                                    encoderToUse = new TiffBitmapEncoder();
                                    ext = ".tiff";
                                }
                                else
                                {
                                    var jpegEnc = new JpegBitmapEncoder();
                                    jpegEnc.QualityLevel = Math.Max(1, Math.Min(100, currentQuality));
                                    encoderToUse = jpegEnc;
                                    ext = ".jpg";
                                }

                                using (var ms = new MemoryStream())
                                {
                                    // If watermarking is requested, use the existing GDI pipeline which also draws the watermark.
                                    // If metadata removal is requested (removeMeta) and watermarking is not used, force a GDI re-encode
                                    // which drops EXIF/PropertyItems by drawing the image into a fresh Bitmap.
                                    if (_settings != null && _settings.WatermarkEnabled && !string.IsNullOrWhiteSpace(_settings.WatermarkFile) && File.Exists(_settings.WatermarkFile))
                                    {
                                        try
                                        {
                                            using (var srcMs = new MemoryStream())
                                            {
                                                var pngEnc = new PngBitmapEncoder();
                                                pngEnc.Frames.Add(BitmapFrame.Create(sourceToEncode));
                                                pngEnc.Save(srcMs);
                                                srcMs.Position = 0;

                                                using (var srcImg = System.Drawing.Image.FromStream(srcMs))
                                                {
                                                    bool needOpaque = encoderToUse is JpegBitmapEncoder || encoderToUse is BmpBitmapEncoder || encoderToUse is TiffBitmapEncoder;
                                                    var pf = needOpaque ? System.Drawing.Imaging.PixelFormat.Format24bppRgb : System.Drawing.Imaging.PixelFormat.Format32bppArgb;
                                                    using (var bmp = new System.Drawing.Bitmap(srcImg.Width, srcImg.Height, pf))
                                                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                                                    {
                                                        g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
                                                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                                                        g.Clear(needOpaque ? System.Drawing.Color.White : System.Drawing.Color.Transparent);
                                                        g.DrawImage(srcImg, 0, 0, srcImg.Width, srcImg.Height);

                                                        try
                                                        {
                                                            using (var wmImg = System.Drawing.Image.FromFile(_settings.WatermarkFile))
                                                            {
                                                                double wmScale = Math.Min(0.25 * srcImg.Width / (double)wmImg.Width, 0.25 * srcImg.Height / (double)wmImg.Height);
                                                                int drawW = Math.Max(1, (int)(wmImg.Width * wmScale));
                                                                int drawH = Math.Max(1, (int)(wmImg.Height * wmScale));
                                                                int x = Math.Max(0, srcImg.Width - drawW - 10);
                                                                int y = Math.Max(0, srcImg.Height - drawH - 10);
                                                                var attr = new System.Drawing.Imaging.ImageAttributes();
                                                                g.DrawImage(wmImg, new System.Drawing.Rectangle(x, y, drawW, drawH), 0, 0, wmImg.Width, wmImg.Height, System.Drawing.GraphicsUnit.Pixel, attr);
                                                            }
                                                        }
                                                        catch { }

                                                        if (encoderToUse is JpegBitmapEncoder)
                                                        {
                                                            var jpgEncoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                                                            var eps = new System.Drawing.Imaging.EncoderParameters(1);
                                                            eps.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)currentQuality);
                                                            bmp.Save(ms, jpgEncoder, eps);
                                                        }
                                                        else if (encoderToUse is BmpBitmapEncoder)
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                                        }
                                                        else if (encoderToUse is TiffBitmapEncoder)
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Tiff);
                                                        }
                                                        else
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            encoderToUse.Frames.Add(BitmapFrame.Create(sourceToEncode));
                                            encoderToUse.Save(ms);
                                        }
                                    }
                                    else if (removeMeta)
                                    {
                                        // Aggressively strip metadata by forcing a re-encode through System.Drawing without copying PropertyItems.
                                        try
                                        {
                                            using (var srcMs = new MemoryStream())
                                            {
                                                var pngEnc = new PngBitmapEncoder();
                                                pngEnc.Frames.Add(BitmapFrame.Create(sourceToEncode));
                                                pngEnc.Save(srcMs);
                                                srcMs.Position = 0;

                                                using (var srcImg = System.Drawing.Image.FromStream(srcMs))
                                                {
                                                    bool needOpaque = encoderToUse is JpegBitmapEncoder || encoderToUse is BmpBitmapEncoder || encoderToUse is TiffBitmapEncoder;
                                                    var pf = needOpaque ? System.Drawing.Imaging.PixelFormat.Format24bppRgb : System.Drawing.Imaging.PixelFormat.Format32bppArgb;
                                                    using (var bmp = new System.Drawing.Bitmap(srcImg.Width, srcImg.Height, pf))
                                                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                                                    {
                                                        g.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
                                                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

                                                        g.Clear(needOpaque ? System.Drawing.Color.White : System.Drawing.Color.Transparent);
                                                        g.DrawImage(srcImg, 0, 0, srcImg.Width, srcImg.Height);

                                                        if (encoderToUse is JpegBitmapEncoder)
                                                        {
                                                            var jpgEncoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                                                            var eps = new System.Drawing.Imaging.EncoderParameters(1);
                                                            eps.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)currentQuality);
                                                            bmp.Save(ms, jpgEncoder, eps);
                                                        }
                                                        else if (encoderToUse is BmpBitmapEncoder)
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                                        }
                                                        else if (encoderToUse is TiffBitmapEncoder)
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Tiff);
                                                        }
                                                        else
                                                        {
                                                            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        catch
                                        {
                                            // fallback
                                            encoderToUse.Frames.Add(BitmapFrame.Create(sourceToEncode));
                                            encoderToUse.Save(ms);
                                        }
                                    }
                                    else
                                    {
                                        encoderToUse.Frames.Add(BitmapFrame.Create(sourceToEncode));
                                        encoderToUse.Save(ms);
                                    }
                                    encodedData = ms.ToArray();
                                }

                                // If Mail-Accept is not requested or we are already under threshold, break
                                if (!mailAccept || encodedData.Length <= mailThresholdBytes)
                                    break;

                                // Otherwise try to reduce quality first (if JPEG), then scale down
                                if (encoderToUse is JpegBitmapEncoder)
                                {
                                    if (currentQuality > 15)
                                    {
                                        currentQuality = Math.Max(10, currentQuality - 15);
                                        continue; // re-encode with lower quality
                                    }
                                }

                                // reduce scale and try again
                                currentScale *= 0.85;
                                if (currentScale < 0.05)
                                    break;
                            }

                            string outName;
                            if (rename)
                            {
                                var safeBase = string.IsNullOrWhiteSpace(renameTxt) ? "image" : renameTxt;
                                outName = System.IO.Path.Combine(destpath, $"{safeBase}_{i + 1}{ext}");
                            }
                            else
                            {
                                var baseName = System.IO.Path.GetFileNameWithoutExtension(item.FileName);
                                outName = System.IO.Path.Combine(destpath, baseName + ext);
                            }

                            var outFileNameOnly = System.IO.Path.GetFileName(outName);
                            if (encodedData == null)
                                throw new Exception("Fehler beim Kodieren der Datei.");

                            if (zipAfter)
                            {
                                lock (createdFilesData)
                                {
                                    // Store encoded bytes for zipping later
                                    createdFilesData.Add((outFileNameOnly, encodedData));
                                }
                            }
                            else
                            {
                                // Write encoded bytes to disk
                                File.WriteAllBytes(outName, encodedData);

                                // remember created file for optional zipping
                                lock (createdFiles)
                                {
                                    createdFiles.Add(outName);
                                }
                            }
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.Invoke(() => WpfMsg.Show("Fehler bei Datei " + item.FileName + ": " + ex.Message));
                    }

                    // update progress
                    var progress = (int)((i + 1) * 100.0 / items.Count);
                    Dispatcher.Invoke(() =>
                    {
                        statusProgressBar.Value = progress;
                    });
                }
            });

            // Finished
            statusProgressBar.Value = 100;
            statusProgressBar.Foreground = new SolidColorBrush(Colors.Green);
            statusText.Text = "Fertig";

            if (zipAfter && (createdFiles.Count > 0 || createdFilesData.Count > 0))
            {
                try
                {
                    var zipName = System.IO.Path.Combine(destpath, "converted_images.zip");
                    // If exists, append timestamp
                    if (File.Exists(zipName))
                    {
                        var time = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        zipName = System.IO.Path.Combine(destpath, $"converted_images_{time}.zip");
                    }

                    // Create zip
                    using (var zip = ZipFile.Open(zipName, ZipArchiveMode.Create))
                    {
                        // files written to disk
                        foreach (var f in createdFiles)
                        {
                            try
                            {
                                zip.CreateEntryFromFile(f, System.IO.Path.GetFileName(f));
                            }
                            catch { }
                        }

                        // files in memory (when zipAfter was selected)
                        lock (createdFilesData)
                        {
                            foreach (var entry in createdFilesData)
                            {
                                try
                                {
                                    var ze = zip.CreateEntry(entry.name);
                                    using (var zs = ze.Open())
                                    using (var ms = new MemoryStream(entry.data))
                                    {
                                        ms.CopyTo(zs);
                                    }
                                }
                                catch { }
                            }
                        }
                    }

                    statusText.Text = "Fertig - ZIP erstellt: " + System.IO.Path.GetFileName(zipName);
                }
                catch (Exception ex)
                {
                    statusText.Text = "Fertig - ZIP fehlgeschlagen: " + ex.Message;
                }
            }
            startButton.IsEnabled = true;
            selectAllCheckBox.IsEnabled = true;
        }

    }
}