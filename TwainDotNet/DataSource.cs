using System;
using System.Collections.Generic;
using System.Text;
using TwainDotNet.TwainNative;

namespace TwainDotNet
{
    public class DataSource : IDisposable
    {
        const ushort TWSX_NATIVE = 1;
        const ushort TWSX_MEMORY = 2;

        Identity _applicationId;
        IWindowsMessageHook _messageHook;

        public DataSource(Identity applicationId, Identity sourceId, IWindowsMessageHook messageHook)
        {
            _applicationId = applicationId;
            SourceId = sourceId.Clone();
            _messageHook = messageHook;
        }

        ~DataSource()
        {
            Dispose(false);
        }

        public Identity SourceId { get; private set; }

        public void NegotiateTransferCount(ScanSettings scanSettings)
        {
            try
            {
                scanSettings.TransferCount = Capability.SetCapability(
                        Capabilities.XferCount,
                        scanSettings.TransferCount,
                        _applicationId,
                        SourceId);
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set TransferCount");
                // Do nothing if the data source does not support the requested capability
            }
        }

        public void NegotiateFeeder(ScanSettings scanSettings)
        {

            try
            {
                if (scanSettings.UseDocumentFeeder.HasValue)
                {
                    Capability.SetCapability(Capabilities.FeederEnabled, scanSettings.UseDocumentFeeder.Value, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set UseDocumentFeeder");
                // Do nothing if the data source does not support the requested capability
            }

            try
            {
                if (scanSettings.UseAutoFeeder.HasValue)
                {
                    Capability.SetCapability(Capabilities.AutoFeed, scanSettings.UseAutoFeeder == true && scanSettings.UseDocumentFeeder == true, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set UseAutoFeeder");
                // Do nothing if the data source does not support the requested capability
            }

            try
            {
                if (scanSettings.UseAutoScanCache.HasValue)
                {
                    Capability.SetCapability(Capabilities.AutoScan, scanSettings.UseAutoScanCache.Value, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set UseAutoScanCache");
                // Do nothing if the data source does not support the requested capability
            }

        }

        public PixelType GetPixelType(ScanSettings scanSettings)
        {
            switch (scanSettings.Resolution.ColourSetting)
            {
                case ColourSetting.BlackAndWhite:
                    return PixelType.BlackAndWhite;

                case ColourSetting.GreyScale:
                    return PixelType.Grey;

                case ColourSetting.Colour:
                    return PixelType.Rgb;
            }

            throw new NotImplementedException();
        }

        public short GetBitDepth(ScanSettings scanSettings)
        {
            switch (scanSettings.Resolution.ColourSetting)
            {
                case ColourSetting.BlackAndWhite:
                    return 1;

                case ColourSetting.GreyScale:
                    return 8;

                case ColourSetting.Colour:
                    return 16;
            }

            throw new NotImplementedException();
        }

        public bool PaperDetectable
        {
            get
            {
                try
                {
                    return Capability.GetBoolCapability(Capabilities.FeederLoaded, _applicationId, SourceId);
                }
                catch
                {
                    return false;
                }
            }
        }

        public bool SupportsDuplex
        {
            get
            {
                try
                {
                    var cap = new Capability(Capabilities.Duplex, TwainType.Int16, _applicationId, SourceId);
                    return ((Duplex)cap.GetBasicValue().Int16Value) != Duplex.None;
                }
                catch
                {
                    return false;
                }
            }
        }

        public void NegotiateColour(ScanSettings scanSettings)
        {
            try
            {
                Capability.SetBasicCapability(Capabilities.IPixelType, (ushort)GetPixelType(scanSettings), TwainType.UInt16, _applicationId, SourceId);
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set IPixelType");
                // Do nothing if the data source does not support the requested capability
            }

            // TODO: Also set this for colour scanning
            try
            {
                if (scanSettings.Resolution.ColourSetting != ColourSetting.Colour)
                {
                    Capability.SetCapability(Capabilities.BitDepth, GetBitDepth(scanSettings), _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set BitDepth");
                // Do nothing if the data source does not support the requested capability
            }

        }

        public void NegotiateResolution(ScanSettings scanSettings)
        {
            try
            {
                if (scanSettings.Resolution.Dpi.HasValue)
                {
                    int dpi = scanSettings.Resolution.Dpi.Value;
                    Capability.SetBasicCapability(Capabilities.XResolution, dpi, TwainType.Fix32, _applicationId, SourceId);
                    Capability.SetBasicCapability(Capabilities.YResolution, dpi, TwainType.Fix32, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Resolution");
                // Do nothing if the data source does not support the requested capability
            }
        }

        public void NegotiateDuplex(ScanSettings scanSettings)
        {
            try
            {
                if (scanSettings.UseDuplex.HasValue && SupportsDuplex)
                {
                    Capability.SetCapability(Capabilities.DuplexEnabled, scanSettings.UseDuplex.Value, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set UseDuplex");
                // Do nothing if the data source does not support the requested capability
            }
        }

        public void NegotiateOrientation(ScanSettings scanSettings)
        {
            // Set orientation (default is portrait)
            try
            {
                var cap = new Capability(Capabilities.Orientation, TwainType.Int16, _applicationId, SourceId);
                if ((Orientation)cap.GetBasicValue().Int16Value != Orientation.Default)
                {
                    Capability.SetBasicCapability(Capabilities.Orientation, (ushort)scanSettings.Page.Orientation, TwainType.UInt16, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Orientation");
                // Do nothing if the data source does not support the requested capability
            }
        }

        /// <summary>
        /// Negotiates the size of the page.
        /// </summary>
        /// <param name="scanSettings">The scan settings.</param>
        public void NegotiatePageSize(ScanSettings scanSettings)
        {
            try
            {
                var cap = new Capability(Capabilities.Supportedsizes, TwainType.Int16, _applicationId, SourceId);
                if ((PageType)cap.GetBasicValue().Int16Value != PageType.UsLetter)
                {
                    Capability.SetBasicCapability(Capabilities.Supportedsizes, (ushort)scanSettings.Page.Size, TwainType.UInt16, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Page.Size");
                // Do nothing if the data source does not support the requested capability
            }
        }

        /// <summary>
        /// Negotiates the automatic rotation capability.
        /// </summary>
        /// <param name="scanSettings">The scan settings.</param>
        public void NegotiateAutomaticRotate(ScanSettings scanSettings)
        {
            try
            {
                if (scanSettings.Rotation.AutomaticRotate)
                {
                    Capability.SetCapability(Capabilities.Automaticrotate, true, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Automaticrotate");
                // Do nothing if the data source does not support the requested capability
            }
        }

        /// <summary>
        /// Negotiates the automatic border detection capability.
        /// </summary>
        /// <param name="scanSettings">The scan settings.</param>
        public void NegotiateAutomaticBorderDetection(ScanSettings scanSettings)
        {
            try
            {
                if (scanSettings.Rotation.AutomaticBorderDetection)
                {
                    Capability.SetCapability(Capabilities.Automaticborderdetection, true, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Automaticborderdetection");
                // Do nothing if the data source does not support the requested capability
            }
        }

        /// <summary>
        /// Negotiates the indicator.
        /// </summary>
        /// <param name="scanSettings">The scan settings.</param>
        public void NegotiateProgressIndicator(ScanSettings scanSettings)
        {
            try
            {
                if (scanSettings.ShowProgressIndicatorUI.HasValue)
                {
                    Capability.SetCapability(Capabilities.Indicators, scanSettings.ShowProgressIndicatorUI.Value, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set ProgressIndicator");
                // Do nothing if the data source does not support the requested capability
            }
        }

        /// <summary>
        /// 先open data source進入state 4，然後才做Capability Negotiation
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="useMemoryMode"></param>
        /// <returns></returns>
        public bool Open(ScanSettings settings, bool useMemoryMode)
        {
            OpenSource();

            if (settings.AbortWhenNoPaperDetectable && !PaperDetectable)
                throw new FeederEmptyException("Feeder is empty.");

            // This is twain state 4,
            // the only state wherein capabilities can be set or reset.
            // Set whether or not to show progress window
            NegotiateProgressIndicator(settings);
            NegotiateTransferCount(settings);
            NegotiateFeeder(settings);
            NegotiateDuplex(settings);

            if (settings.UseDocumentFeeder == true &&
                settings.Page != null)
            {
                NegotiatePageSize(settings);
                NegotiateOrientation(settings);
            }

            if (settings.Area != null)
            {
                NegotiateArea(settings);
            }

            if (settings.Resolution != null)
            {
                NegotiateColour(settings);
                NegotiateResolution(settings);
            }

            // Configure automatic rotation and image border detection
            if (settings.Rotation != null)
            {
                NegotiateAutomaticRotate(settings);
                NegotiateAutomaticBorderDetection(settings);
            }

            if (useMemoryMode)
            {
                NegotiateBufferedMode(TWSX_MEMORY);
            }

            NegotiateTest();
            NegotiateContrast(settings.Contrast);
            NegotiateBrightness(120);

            // Go from twain state 4 to 5
            return Enable(settings);
        }

        private bool NegotiateTest()
        {
            try
            {
                //var cap = new Capability(Capabilities.Brightness, TwainType.Fix32, _applicationId, SourceId);
                //int irtn = cap.GetBasicValue().Int32Value;
                //Capability.SetCapability(Capabilities.Brightness, 100, _applicationId, SourceId);
                //Capability.SetBasicCapability(Capabilities.Contrast, 300, TwainType.Fix32, _applicationId, SourceId);
                //if (iRtn > 0)
                {
                    bool brtn = Capability.GetBoolCapability(Capabilities.Autobright, _applicationId, SourceId);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return false;
            }
            return true;
        }

        /// <summary>
        /// A8專用的校正，需要先select一個data source且先Open source。
        /// </summary>
        public bool CalibrateA8()
        {
            /*
             * 進行Calibration的動作。這不是TWAIN spec裡的東西，需要廠商提供。
             * 來自XTWainAVISION.cpp，XTWainAVISION::DoCalibration()
             * BOOL bRtna = SetCapability(CAP_INDICATORS, FALSE, TWTY_BOOL);
	         * BOOL bRtn = SetCapability(0x9259, TRUE, TWTY_BOOL);
             */
            bool bReturn = false;
            try
            {
                Capability.SetCapability(Capabilities.Indicators, false, _applicationId, SourceId);
                //int iReturn = Capability.SetBasicCapability(Capabilities.Indicators, 0, TwainType.Bool, _applicationId, SourceId);
                Capability.SetCapability(Capabilities.A8_Calibrate, true, _applicationId, SourceId);
                /*int iReturn = Capability.SetBasicCapability(Capabilities.A8_Calibrate, 1, TwainType.Bool, _applicationId, SourceId);
                if (iReturn > 0)
                    bReturn = true;*/
            }
            catch (Exception ex)
            {
                return false;
            }
            return bReturn;
        }

        /// <summary>
        /// A8是否需要校正，需要先select一個data source且先Open source。
        /// </summary>
        public bool A8NeedCalibrate()
        {
            //bNeed = FALSE;
            //float fVal;
            //BOOL bRtn = GetCapability(0x9259, fVal, MSG_GETCURRENT); ///Carpe: id 仍未確認 
            //bNeed = (BOOL)(fVal + 0.00001);

            Capability cap = new Capability(Capabilities.A8_Calibrate, TwainType.Int32, _applicationId, SourceId);
            int iReturn = cap.GetBasicValue().Int32Value;
            if (iReturn > 0)
                return true;
            else
                return false;
        }

        private bool NegotiateContrast(int contrast)
        {
            try
            {
                Capability cap = new Capability(Capabilities.Contrast, TwainType.Fix32, _applicationId, SourceId);
                int oldContrast = cap.GetBasicValue().Int32Value;
                if (oldContrast != contrast)
                {
                    int iRtn = Capability.SetBasicCapability(Capabilities.Contrast, contrast, TwainType.Fix32, _applicationId, SourceId);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Contrast");
                return false;
            }
            return true;
        }

        private bool NegotiateBrightness(int brightness)
        {
            try
            {
                Capability cap = new Capability(Capabilities.Brightness, TwainType.Fix32, _applicationId, SourceId);
                int oldContrast = cap.GetBasicValue().Int32Value;
                if (oldContrast != brightness)
                {
                    int iRtn = Capability.SetBasicCapability(Capabilities.Brightness, brightness, TwainType.Fix32, _applicationId, SourceId);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Brightness");
                return false;
            }
            return true;
        }

        private bool NegotiateBufferedMode(ushort scanMode)
        {
            try
            {                
                int iRtn = Capability.SetBasicCapability(Capabilities.IXferMech, scanMode, TwainType.UInt16, _applicationId, SourceId);
                //Console.WriteLine(iRtn);
            }
            catch (Exception ex)
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set BufferedMode");
                // Do nothing if the data source does not support the requested capability
            }
            return true;
        }

        private bool NegotiateArea(ScanSettings scanSettings)
        {
            var area = scanSettings.Area;

            if (area == null)
            {
                return false;
            }

            try
            {
                var cap = new Capability(Capabilities.IUnits, TwainType.UInt16, _applicationId, SourceId);
                if ((Units)cap.GetBasicValue().Int16Value != area.Units)
                {
                    Capability.SetBasicCapability(Capabilities.IUnits, (int)area.Units, TwainType.UInt16, _applicationId, SourceId);
                }
            }
            catch
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set Units");
                // Do nothing if the data source does not support the requested capability
            }

            var imageLayout = new ImageLayout
            {
                Frame = new Frame
                {
                    Left = new Fix32(area.Left),
                    Top = new Fix32(area.Top),
                    Right = new Fix32(area.Right),
                    Bottom = new Fix32(area.Bottom)
                }
            };

            var result = Twain32Native.DsImageLayout(
                _applicationId,
                SourceId,
                DataGroup.Image,
                DataArgumentType.ImageLayout,
                Message.Set,
                imageLayout);

            if (result != TwainResult.Success)
            {
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "DataSource failed to set ImageLayout");
                // note: 我這裡總是失敗，但是實際上image layout是可以設定成功的，由掃描結果不同來推測。
                //throw new TwainException("DsImageLayout.GetDefault error", result);
            }

            return true;
        }

        /// <summary>
        /// 打開DS，進入State 4以溝通Capabilities
        /// </summary>
        /// <exception cref="DeviceOpenExcetion">當DS被其他人占用的時候可能會擲回，或twain DLL沒有錯誤，但是回傳值失敗的話就會擲回</exception>
        public void OpenSource()
        {
            TwainResult result = TwainResult.NotDSEvent;
            try
            {
                result = Twain32Native.DsmIdentity(
                       _applicationId,
                       IntPtr.Zero,
                       DataGroup.Control,
                       DataArgumentType.Identity,
                       Message.OpenDS,
                       SourceId);
            }
            catch (Exception ex)
            {
                throw new DeviceOpenExcetion(ex.Message);
            }

            if (result != TwainResult.Success)
            {
                throw new DeviceOpenExcetion("Error opening data source", result);
            }
            Logger.WriteLog(LOG_LEVEL.LL_NORMAL_LOG, string.Format("DataSource \"{0}\" successfully opened.", this.SourceId.ProductName));
        }

        /// <summary>
        /// 接下來就交給source控制這個流程。MSG_XFERREADY, MSG_CLOSEDSREQ, or MSG_CLOSEDSOK messages會傳回來，其中MSG_XFERREADY代表source準備好從state 5→6。
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public bool Enable(ScanSettings settings)
        {
            UserInterface ui = new UserInterface();
            ui.ShowUI = (settings.ShowTwainUI ? TwainBool.True : TwainBool.False);
            ui.ModalUI = TwainBool.False;
            ui.ParentHand = _messageHook.WindowHandle;

            var result = Twain32Native.DsUserInterface(
                _applicationId,
                SourceId,
                DataGroup.Control,
                DataArgumentType.UserInterface,
                Message.EnableDS,
                ui);

            if (result != TwainResult.Success)
            {
                Console.WriteLine("Enable() error! return=" + result);
                Logger.WriteLog(LOG_LEVEL.LL_SUB_FUNC, "Enable() error! return=" + result);
                Dispose();
                return false;
            }
            return true;
        }

        /// <summary>
        /// Get default data source.
        /// </summary>
        /// <param name="applicationId"></param>
        /// <param name="messageHook"></param>
        /// <returns></returns>
        public static DataSource GetDefault(Identity applicationId, IWindowsMessageHook messageHook)
        {
            var defaultSourceId = new Identity();

            // Attempt to get information about the system default source
            var result = Twain32Native.DsmIdentity(
                applicationId,
                IntPtr.Zero,
                DataGroup.Control,
                DataArgumentType.Identity,
                Message.GetDefault,
                defaultSourceId);

            if (result != TwainResult.Success)
            {
                ConditionCode status = DataSourceManager.GetConditionCode(applicationId, null);
                Logger.WriteLog(LOG_LEVEL.LL_SERIOUS_ERROR, "Error getting information about the default source: " + result);
                throw new TwainException("Error getting information about the default source: " + result, result, status);
            }

            return new DataSource(applicationId, defaultSourceId, messageHook);
        }

        public static DataSource UserSelected(Identity applicationId, IWindowsMessageHook messageHook)
        {
            var defaultSourceId = new Identity();

            // Show the TWAIN interface to allow the user to select a source
            Twain32Native.DsmIdentity(
                applicationId,
                IntPtr.Zero,
                DataGroup.Control,
                DataArgumentType.Identity,
                Message.UserSelect,
                defaultSourceId);

            return new DataSource(applicationId, defaultSourceId, messageHook);
        }

        public static List<DataSource> GetAllSources(Identity applicationId, IWindowsMessageHook messageHook)
        {
            var sources = new List<DataSource>();
            Identity id = new Identity();

            // Get the first source
            var result = Twain32Native.DsmIdentity(
                applicationId,
                IntPtr.Zero,
                DataGroup.Control,
                DataArgumentType.Identity,
                Message.GetFirst,
                id);

            if (result == TwainResult.EndOfList)
            {
                return sources;
            }
            else if (result != TwainResult.Success)
            {
                throw new TwainException("Error getting first source.", result);
            }
            else
            {
                sources.Add(new DataSource(applicationId, id, messageHook));
            }

            while (true)
            {
                // Get the next source
                result = Twain32Native.DsmIdentity(
                    applicationId,
                    IntPtr.Zero,
                    DataGroup.Control,
                    DataArgumentType.Identity,
                    Message.GetNext,
                    id);

                if (result == TwainResult.EndOfList)
                {
                    break;
                }
                else if (result != TwainResult.Success)
                {
                    throw new TwainException("Error enumerating sources.", result);
                }

                sources.Add(new DataSource(applicationId, id, messageHook));
            }

            return sources;
        }

        public static DataSource GetSource(string sourceProductName, Identity applicationId, IWindowsMessageHook messageHook)
        {
            // A little slower than it could be, if enumerating unnecessary sources is slow. But less code duplication.
            foreach (var source in GetAllSources(applicationId, messageHook))
            {
                if (sourceProductName.Equals(source.SourceId.ProductName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return source;
                }
            }

            return null;
        }


        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected void Dispose(bool disposing)
        {
            if (disposing)
            {
                Close();
            }
        }

        public void Close()
        {
            if (SourceId.Id != 0)
            {
                UserInterface userInterface = new UserInterface();

                TwainResult result = Twain32Native.DsUserInterface(
                    _applicationId,
                    SourceId,
                    DataGroup.Control,
                    DataArgumentType.UserInterface,
                    Message.DisableDS,
                    userInterface);

                if (result != TwainResult.Failure)
                {
                    result = Twain32Native.DsmIdentity(
                        _applicationId,
                        IntPtr.Zero,
                        DataGroup.Control,
                        DataArgumentType.Identity,
                        Message.CloseDS,
                        SourceId);
                }
            }
        }
    }
}
