using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Switcheroo.Core;

namespace Switcheroo
{
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
        private static BitmapImage ConvertIconToBitmap(System.Drawing.Icon sysicon)
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

    public class AppWindowViewModel : ListItemInfo, INotifyPropertyChanged, IWindowText
    {
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

        public event PropertyChangedEventHandler PropertyChangedHwnd;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChangedHwnd?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public new string FormattedTitle
        {
            get => base.FormattedTitle;
            set
            {
                base.FormattedTitle = value;
                NotifyOfPropertyChange(() => FormattedTitle);
            }
        }

        public new string FormattedSubTitle
        {
            get => base.FormattedSubTitle;
            set
            {
                base.FormattedSubTitle = value;
                NotifyOfPropertyChange(() => FormattedSubTitle);
            }
        }

        public new BitmapImage ImageSource
        {
            get => base.ImageSource;
            set
            {
                base.ImageSource = value;
                NotifyOfPropertyChange(() => ImageSource);
            }
        }

        private bool _isBeingClosed = false;

        public bool IsBeingClosed
        {
            get { return _isBeingClosed; }
            set
            {
                _isBeingClosed = value;
                NotifyOfPropertyChange(() => IsBeingClosed);
            }
        }

        #endregion

        #region INotifyPropertyChanged Members

        public event PropertyChangedEventHandler PropertyChanged;

        private void NotifyOfPropertyChange<T>(Expression<Func<T>> property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(GetPropertyName(property)));
        }

        private string GetPropertyName<T>(Expression<Func<T>> property)
        {
            var lambda = (LambdaExpression) property;

            MemberExpression memberExpression;
            if (lambda.Body is UnaryExpression unaryExpression)
            {
                memberExpression = (MemberExpression)unaryExpression.Operand;
            }
            else
            {
                memberExpression = (MemberExpression)lambda.Body;
            }

            return memberExpression.Member.Name;
        }
        #endregion
    }

    public class ListItemInfo 
    {
        public ListItemInfo() { }
        public ListItemInfo(string title, string subTitle, BitmapImage imageSource)
        {
            FormattedTitle = title;
            FormattedSubTitle = subTitle;
            ImageSource = imageSource;
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
            get { return _bitmapImage;}
            set { _bitmapImage = value; }
        }
    }
}