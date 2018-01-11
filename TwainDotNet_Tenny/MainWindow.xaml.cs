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
using System.Drawing;
using TwainDotNet;
using TwainDotNet.TwainNative;
using TwainDotNet.Win32;
using System.Windows.Interop;
using System.IO;
using System.Threading;

namespace TwainDotNet_Tenny
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private Twain _twain;
        private ScanSettings _settings;
        private Bitmap resultImage;
        private int imageCount;
        private string scannerName;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //_twain = new Twain(new WpfWindowMessageHook(this));
            //_twain.TransferImage += _twain_TransferImage;
            //_twain.ScanningComplete += delegate
            //{
            //    IsEnabled = true;
            //};

            //IList<string> sourceList = _twain.SourceNames;
            //ManualSource.ItemsSource = sourceList;
            //if (sourceList != null && sourceList.Count > 0)
            //    ManualSource.SelectedItem = sourceList[0];
        }

        private void _twain_TransferImage(object sender, TransferImageEventArgs e)
        {
            //IsEnabled = true;
            if (e.Image != null)
            {
                resultImage = e.Image;
                string savePath = @"C:\Users\Tenny\Pictures\TwainTest\testBufferPic_";
                savePath += imageCount.ToString() + @".bmp";
                resultImage.Save(savePath, System.Drawing.Imaging.ImageFormat.Bmp);
                fileName.Content = savePath;

                IntPtr hbitmap = new Bitmap(e.Image).GetHbitmap();
                MainImage.Source = Imaging.CreateBitmapSourceFromHBitmap(
                        hbitmap,
                        IntPtr.Zero,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());
                Gdi32Native.DeleteObject(hbitmap);

                imageCount++;
            }
        }

        private void btnInit_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("Scan...");
            fileName.Content = string.Empty;
            progressBar.Visibility = Visibility.Visible;

            Thread thread = new Thread(() =>
            {
                ScanWindow sw = new ScanWindow();
                sw.TransferTempImage += Sw_TransferTempImage;
                sw.TransferCompleteImage += Sw_TransferCompleteImage;
                sw.ShowDialog();
            });
            thread.TrySetApartmentState(ApartmentState.STA);
            thread.Start();
            return;

            //_settings = new ScanSettings();
            //_settings.ShouldTransferAllPages = true;
            //_settings.ShowProgressIndicatorUI = false;
            //_settings.Resolution = ResolutionSettings.ColourPhotocopier;
            imageCount = 0;

            _settings = new ScanSettings
            {
                UseDocumentFeeder = false,
                ShowTwainUI = false,
                ShowProgressIndicatorUI = false,
                UseDuplex = false,
                Resolution = ResolutionSettings.Photocopier,
                Area = null,
                ShouldTransferAllPages = true,
                Rotation = new RotationSettings
                {
                    AutomaticRotate = false,
                    AutomaticBorderDetection = false
                }
            };

            scannerName = ManualSource.SelectedItem.ToString();

            try
            {
                _twain.SelectSource(scannerName);
                _twain.StartScanning(_settings);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            Console.WriteLine("Scan done.");
        }

        private void Sw_TransferCompleteImage(object sender, TransferImageEventArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    IntPtr hbitmap = new Bitmap(e.Image).GetHbitmap();
                    MainImage.Source = Imaging.CreateBitmapSourceFromHBitmap(
                            hbitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromRotation(Rotation.Rotate270));
                    Gdi32Native.DeleteObject(hbitmap);
                    fileName.Content = "complete!";
                    progressBar.Visibility = Visibility.Hidden;
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void Sw_TransferTempImage(object sender, TransferImageEventArgs e)
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    IntPtr hbitmap = new Bitmap(e.Image).GetHbitmap();
                    MainImage.Source = Imaging.CreateBitmapSourceFromHBitmap(
                            hbitmap,
                            IntPtr.Zero,
                            Int32Rect.Empty,
                            BitmapSizeOptions.FromRotation(Rotation.Rotate270));
                    Gdi32Native.DeleteObject(hbitmap);
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _twain?.Dispose();
            Console.WriteLine("Window closing");
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            MainImage.Source = null;
            string savePath = @"C:\Users\Tenny\Pictures\TwainTest\";
            try
            {
                DirectoryInfo di = new DirectoryInfo(savePath);
                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }
            }
            catch { }
        }

        private void btnPaperOn_Click(object sender, RoutedEventArgs e)
        {
            bool isPaperOn = false;
            Twain twain = null;
            try
            {
                twain = new Twain(new WpfWindowMessageHook(this));
                twain.SelectSource("A8 ColorScanner PP");
                isPaperOn = twain.IsPaperOn();
                twain.Dispose();
            }
            catch (TwainException ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }

            if (isPaperOn)
            {
                MessageBox.Show("有紙");
            }
            else
            {
                MessageBox.Show("沒紙");
            }            
        }

        private void btnNeedCalibrate_Click(object sender, RoutedEventArgs e)
        {
            bool isPaperOn = false;
            Twain twain = null;
            try
            {
                twain = new Twain(new WpfWindowMessageHook(this));
                twain.SelectSource("A8 ColorScanner PP");
                isPaperOn = twain.A8NeedCalibrate();
                twain.Dispose();
            }
            catch (TwainException ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }

            if (isPaperOn)
            {
                MessageBox.Show("需要校正");
            }
            else
            {
                MessageBox.Show("不用校正");
            }
        }

        private void btnCalibrate_Click(object sender, RoutedEventArgs e)
        {
            bool isPaperOn = false;
            Twain twain = null;
            try
            {
                twain = new Twain(new WpfWindowMessageHook(this));
                twain.SelectSource("A8 ColorScanner PP");
                isPaperOn = twain.CalibrateA8();
                twain.Dispose();
            }
            catch (TwainException ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                isPaperOn = false;
                twain?.Dispose();
            }
        }
    }
}
