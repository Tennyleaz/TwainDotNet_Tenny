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
using System.Windows.Shapes;
using TwainDotNet;
using TwainDotNet.TwainNative;
using TwainDotNet.Win32;
using System.Windows.Interop;
using System.Drawing;
using System.Windows.Threading;
using System.Drawing.Imaging;

namespace TwainDotNet_Tenny
{
    /// <summary>
    /// ScanWindow.xaml 的互動邏輯
    /// </summary>
    public partial class ScanWindow : Window
    {
        private Twain _twain;
        private ScanSettings _settings;
        private int imageCount = 0;
        private string _scannerName;
        private Bitmap resultImage;
        private IList<string> sourceList;
        private DispatcherTimer timer;

        public delegate void TransferTempImageHandler(object sender, TransferImageEventArgs e);  //定義Event handler
        public event TransferTempImageHandler TransferTempImage;  //做一份實例
        public event TransferTempImageHandler TransferCompleteImage;  //做一份實例

        public ScanWindow()
        {
            InitializeComponent();
            Loaded += ScanWindow_Loaded;
            Closing += Window_Closing;           
        }

        public void OnTransferTempImage(Bitmap image, bool continueScanning, float percentage)
        {
            TransferImageEventArgs e = new TransferImageEventArgs(image, continueScanning, percentage);
            TransferTempImage?.Invoke(this, e);
        }

        public void OnTransferComplete(Bitmap image)
        {
            TransferImageEventArgs e = new TransferImageEventArgs(image, false, 1);
            TransferCompleteImage?.Invoke(this, e);
        }

        private Bitmap SwapRedAndBlueChannels(Bitmap bitmap)
        {
            var imageAttr = new ImageAttributes();
            imageAttr.SetColorMatrix(new ColorMatrix(
                                         new[]
                                             {
                                                 new[] {0.0F, 0.0F, 1.0F, 0.0F, 0.0F},
                                                 new[] {0.0F, 1.0F, 0.0F, 0.0F, 0.0F},
                                                 new[] {1.0F, 0.0F, 0.0F, 0.0F, 0.0F},
                                                 new[] {0.0F, 0.0F, 0.0F, 1.0F, 0.0F},
                                                 new[] {0.0F, 0.0F, 0.0F, 0.0F, 1.0F}
                                             }
                                         ));
            var temp = new Bitmap(bitmap.Width, bitmap.Height);
            GraphicsUnit pixel = GraphicsUnit.Pixel;
            using (Graphics g = Graphics.FromImage(temp))
            {
                g.DrawImage(bitmap, System.Drawing.Rectangle.Round(bitmap.GetBounds(ref pixel)), 0, 0, bitmap.Width, bitmap.Height,
                            GraphicsUnit.Pixel, imageAttr);
            }

            return temp;
        }

        private void _twain_TransferImage(object sender, TransferImageEventArgs e)
        {
            //IsEnabled = true;
            if (e.Image != null)
            {
                resultImage = e.Image;
                resultImage = SwapRedAndBlueChannels(resultImage);
                string savePath = @"C:\Users\Tenny\Pictures\TwainTest\testBufferPic_";
                savePath += imageCount.ToString() + @".bmp";
                resultImage.Save(savePath, System.Drawing.Imaging.ImageFormat.Bmp);
                OnTransferTempImage(resultImage, e.ContinueScanning, e.PercentComplete);
                //fileName.Content = savePath;

                //IntPtr hbitmap = new Bitmap(e.Image).GetHbitmap();
                //MainImage.Source = Imaging.CreateBitmapSourceFromHBitmap(
                //        hbitmap,
                //        IntPtr.Zero,
                //        Int32Rect.Empty,
                //        BitmapSizeOptions.FromEmptyOptions());
                //Gdi32Native.DeleteObject(hbitmap);

                imageCount++;
            }
        }

        private void ScanWindow_Loaded(object sender, RoutedEventArgs e)
        {
            _twain = new Twain(new WpfWindowMessageHook(this));
            _twain.TransferImage += _twain_TransferImage;
            _twain.ScanningComplete += delegate
            {
                OnTransferComplete(resultImage);
                IsEnabled = true;
                this.Close();
            };
            sourceList = _twain.SourceNames;

            timer = new DispatcherTimer();
            timer.Tick += Timer_Tick;
            timer.Interval = new TimeSpan(0, 0, 0, 0, 500);
            timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            timer.Stop();

            imageCount = 0;
            _settings = new ScanSettings
            {
                UseDocumentFeeder = true,
                ShowTwainUI = false,
                ShowProgressIndicatorUI = false,
                UseDuplex = false,
                Resolution = ResolutionSettings.ColourPhotocopier,  // 顏色
                /*Area = null,*/
                ShouldTransferAllPages = true,
                Rotation = new RotationSettings
                {
                    AutomaticRotate = false,
                    AutomaticBorderDetection = false  // A8好像不支援這個
                }
            };
            _settings.Page = new PageSettings();  // page參數好像沒什麼用
            _settings.Page.Size = PageType.None;
            //_settings.Page.Orientation = TwainDotNet.TwainNative.Orientation.Auto;
            _settings.Area = new AreaSettings(Units.Inches, 0, 0, 3.7f, 2.1f);  // 設定掃描紙張的邊界
            _settings.AbortWhenNoPaperDetectable = true;  // 要偵測有沒有紙
            _settings.Contrast = 300;

            try
            {
                _scannerName = /*"WIA-A8 ColorScanner PP"; */"A8 ColorScanner PP"; //sourceList.Last()*/;
                _twain.SelectSource(_scannerName);
                _twain.StartScanning(_settings);  // 做open→設定各種CAPABILITY→EnableDS→等回傳圖
                /// Once the Source is enabled via the DG_CONTROL / DAT_USERINTERFACE/ MSG_ENABLEDS operation, 
                /// all events that enter the application’s main event loop must be immediately forwarded to the Source.
                /// The Source, not the application, controls the transition from State 5 to State 6.
                /// - spec pdf p.3-60
            }
            catch (FeederEmptyException)
            {
                MessageBox.Show("沒有紙");
                Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                Close();
            }            
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _twain?.Dispose();
            Console.WriteLine("Window closing");
        }


    }
}
