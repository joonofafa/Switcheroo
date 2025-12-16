using Switcheroo.Core;
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Media.Imaging;

namespace Switcheroo {
    public static class IconHelper {
        // 창 핸들로부터 아이콘을 가져오기 위한 상수와 메소드 선언
        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_LARGEICON = 0x0; // 대아이콘
        private const uint SHGFI_SMALLICON = 0x1; // 작은 아이콘

        [DllImport("shell32.dll")]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct SHFILEINFO {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        // 창 핸들로부터 아이콘을 가져오는 메소드
        public static BitmapImage GetIconFromWindowHandle(IntPtr hWnd)
        {
            SHFILEINFO shinfo = new SHFILEINFO();
            IntPtr hImg = SHGetFileInfo(hWnd.ToString(), 0, ref shinfo, (uint)Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_SMALLICON);

            if (hImg != IntPtr.Zero)
            {
                // Icon을 BitmapImage로 변환
                using (System.Drawing.Icon sysicon = System.Drawing.Icon.FromHandle(shinfo.hIcon))
                {
                    BitmapImage iconBitmap = ConvertIconToBitmap(sysicon);
                    return iconBitmap;
                }
            }

            return null;
        }

        // Icon을 BitmapImage로 변환하는 메소드
        public static BitmapImage ConvertIconToBitmap(System.Drawing.Icon sysicon)
        {
            using (System.Drawing.Bitmap bitmap = sysicon.ToBitmap())
            {
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                System.IO.MemoryStream memoryStream = new System.IO.MemoryStream();
                bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                memoryStream.Seek(0, System.IO.SeekOrigin.Begin);
                bitmapImage.StreamSource = memoryStream;
                bitmapImage.EndInit();
                return bitmapImage;
            }
        }

        public static BitmapImage LoadImage(string filePath)
        {
            try
            {
                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.UriSource = new Uri(filePath);
                bitmapImage.EndInit();
                return bitmapImage;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading image: {ex.Message}");
                return null;
            }
        }
    }

    public class AppWindowViewModel : ListItemInfo, IWindowText {
        public AppWindowViewModel(AppWindow appWindow)
        {
            AppWindow = appWindow;
        }

        public AppWindowViewModel() { }

        public AppWindow AppWindow { get; private set; }

        #region IWindowText Members

        public string WindowTitle
        {
            get { return AppWindow.Title; }
        }

        public string ProcessTitle
        {
            get { return AppWindow.ProcessTitle; }
        }

        #endregion

        #region Bindable properties

        public IntPtr HWnd
        {
            get { return AppWindow.HWnd; }
        }

        public BitmapImage WindowIcon => IconHelper.GetIconFromWindowHandle(HWnd);

        public new string FormattedTitle
        {
            get => base.FormattedTitle;
            set
            {
                base.FormattedTitle = value;
                OnPropertyChanged(nameof(FormattedTitle));
            }
        }

        public new string FormattedSubTitle
        {
            get => base.FormattedSubTitle;
            set
            {
                base.FormattedSubTitle = value;
                OnPropertyChanged(nameof(FormattedSubTitle));
            }
        }

        public new BitmapImage ImageSource
        {
            get => base.ImageSource;
            set
            {
                base.ImageSource = value;
                OnPropertyChanged(nameof(ImageSource));
            }
        }

        private bool _isBeingClosed = false;

        public bool IsBeingClosed
        {
            get { return _isBeingClosed; }
            set
            {
                _isBeingClosed = value;
                OnPropertyChanged(nameof(IsBeingClosed));
            }
        }

        #endregion
    }

    public class ListItemInfo : INotifyPropertyChanged {
        public ListItemInfo() { }
        public ListItemInfo(string title, string subTitle, BitmapImage imageSource, string tagData, bool isUrl)
        {
            FormattedTitle = title;
            FormattedSubTitle = subTitle;
            ImageSource = imageSource;
            TagData = tagData;
            IsUrl = isUrl;
        }

        public ListItemInfo(string title, string subTitle, BitmapImage imageSource, string tagData, bool isUrl, string argument)
        {
            FormattedTitle = title;
            FormattedSubTitle = subTitle;
            ImageSource = imageSource;
            TagData = tagData;
            IsUrl = isUrl;
            Argument = argument;
        }

        private string _formattedTitle;

        public string FormattedTitle
        {
            get { return _formattedTitle; }
            set { _formattedTitle = value; }
        }

        private string _formattedSubTitle;

        public string FormattedSubTitle
        {
            get { return _formattedSubTitle; }
            set { _formattedSubTitle = value; }
        }

        private BitmapImage _bitmapImage = null;
        public BitmapImage ImageSource
        {
            get { return _bitmapImage; }
            set { _bitmapImage = value; }
        }

        private string _tagData;

        public string TagData
        {
            get { return _tagData; }
            set { _tagData = value; }
        }

        private bool _isUrl;
        public bool IsUrl
        {
            get { return _isUrl; }
            set { _isUrl = value; }
        }

        private string _argument;
        public string Argument
        {
            get { return _argument; }
            set { _argument = value; }
        }

        // Icon information for lazy loading
        private string _iconPath;
        public string IconPath
        {
            get { return _iconPath; }
            set { _iconPath = value; }
        }

        private int _iconIndex;
        public int IconIndex
        {
            get { return _iconIndex; }
            set { _iconIndex = value; }
        }

        private bool _isDefaultIcon;
        public bool IsDefaultIcon
        {
            get { return _isDefaultIcon; }
            set { _isDefaultIcon = value; }
        }

        // Flag to track if async loading has started
        private bool _iconLoadingStarted = false;

        // Lazy loading property for ImageSource with async support
        public BitmapImage LazyImageSource
        {
            get
            {
                if (_bitmapImage == null && !string.IsNullOrEmpty(_iconPath) && !_iconLoadingStarted)
                {
                    _iconLoadingStarted = true;
                    // Load icon asynchronously
                    LoadIconAsync();
                }
                return _bitmapImage;
            }
        }

        private async void LoadIconAsync()
        {
            try
            {
                string iconPath = _iconPath;
                int iconIndex = _iconIndex;
                bool isDefaultIcon = _isDefaultIcon;

                // Extract icon bytes on background thread
                byte[] iconBytes = await System.Threading.Tasks.Task.Run(() =>
                {
                    return ExtractIconBytes(iconPath, iconIndex, isDefaultIcon);
                });

                // Convert to BitmapImage on UI thread
                if (iconBytes != null && iconBytes.Length > 0)
                {
                    System.Windows.Application.Current?.Dispatcher?.Invoke(() =>
                    {
                        try
                        {
                            var bitmapImage = new BitmapImage();
                            bitmapImage.BeginInit();
                            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                            bitmapImage.StreamSource = new MemoryStream(iconBytes);
                            bitmapImage.EndInit();
                            bitmapImage.Freeze();
                            
                            _bitmapImage = bitmapImage;
                            OnPropertyChanged(nameof(LazyImageSource));
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error creating BitmapImage: {ex.Message}");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading icon async: {ex.Message}");
            }
        }

        private byte[] ExtractIconBytes(string iconPath, int iconIndex, bool isDefaultIcon)
        {
            try
            {
                IntPtr largeIcon = IntPtr.Zero;
                System.Drawing.Icon icon = null;

                try
                {
                    string pathToUse = iconPath;
                    int indexToUse = iconIndex;

                    // For default icon or exe/dll files
                    if (isDefaultIcon || 
                        iconPath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) ||
                        iconPath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
                    {
                        int result = ExtractIconEx(pathToUse, indexToUse, out largeIcon, out _, 1);
                        if (result > 0 && largeIcon != IntPtr.Zero)
                        {
                            icon = System.Drawing.Icon.FromHandle(largeIcon);
                        }
                    }
                    else if (iconPath.EndsWith(".ico", StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(iconPath))
                        {
                            icon = new System.Drawing.Icon(iconPath);
                        }
                    }
                    else
                    {
                        // Try to extract from file
                        int result = ExtractIconEx(pathToUse, indexToUse, out largeIcon, out _, 1);
                        if (result > 0 && largeIcon != IntPtr.Zero)
                        {
                            icon = System.Drawing.Icon.FromHandle(largeIcon);
                        }
                    }

                    // Fallback to default icon
                    if (icon == null)
                    {
                        int result = ExtractIconEx("C:\\Windows\\System32\\shell32.dll", 2, out largeIcon, out _, 1);
                        if (result > 0 && largeIcon != IntPtr.Zero)
                        {
                            icon = System.Drawing.Icon.FromHandle(largeIcon);
                        }
                    }

                    if (icon != null)
                    {
                        using (var bitmap = icon.ToBitmap())
                        using (var ms = new MemoryStream())
                        {
                            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                            return ms.ToArray();
                        }
                    }
                }
                finally
                {
                    if (icon != null)
                        icon.Dispose();
                    if (largeIcon != IntPtr.Zero)
                        DestroyIcon(largeIcon);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting icon bytes: {ex.Message}");
            }

            return null;
        }

        #region INotifyPropertyChanged
        
        public event PropertyChangedEventHandler PropertyChanged;
        
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        
        #endregion

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr largeIcon, out IntPtr smallIcon, int nIcons);
        
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);
    }
}