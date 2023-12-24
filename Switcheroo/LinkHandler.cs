using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Diagnostics;
using IWshRuntimeLibrary;
using System.Runtime.InteropServices;
using System.Text;
using System.Runtime.InteropServices.ComTypes;

namespace Switcheroo {
    class LinkHandler {

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        extern static int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr largeIcon, out IntPtr smallIcon, int nIcons);
        [DllImport("user32.dll", SetLastError = true)]
        extern static bool DestroyIcon(IntPtr hIcon);

        private List<ListItemInfo> _listItemInfos = new List<ListItemInfo>();

        public void cacheLinks()
        {
            string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            if (Directory.Exists(startMenuPath))
            {
                List<string> lnkFiles = FindLnkFiles(startMenuPath);
                foreach (var file in lnkFiles)
                {
                    ListItemInfo listItemInfo = new ListItemInfo
                    {
                        FormattedTitle = Path.GetFileNameWithoutExtension(file),
                        FormattedSubTitle = Path.GetFileNameWithoutExtension(file),
                        ImageSource = GetIconImageFromShortcut(file),
                        TagData = file,
                    };
                    _listItemInfos.Add(listItemInfo);
                }
            }
        }

        private List<string> FindLnkFiles(string directoryPath)
        {
            DirectoryInfo di = new DirectoryInfo(directoryPath);
            return di.EnumerateFiles("*.lnk", SearchOption.AllDirectories)
                     .Where(fi => !fi.Name.Contains("Uninstall"))
                     .Select(fi => fi.FullName)
                     .ToList();
        }

        private BitmapImage GetIconImageFromShortcut(string shortcutPath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            string iconLocation = shortcut.IconLocation;
            string targetPath = shortcut.TargetPath;

            if (string.IsNullOrEmpty(iconLocation))
            {
                return null;
            }

            string[] iconPathParts = iconLocation.Split(',');
            string iconPath = (iconPathParts[0].Trim().Length > 1) ? iconPathParts[0].Trim() : targetPath;
            int iconIndex = (iconPathParts.Length > 1) ? int.Parse(iconPathParts[1]) : 0;

            try
            {
                return GetIconImageFromExecutable(iconPath, iconIndex);
            }
            catch (ArgumentException)
            {
                return GetIconImageFromExecutable("C:\\Windows\\System32\\shell32.dll", 2);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
        }

        public List<ListItemInfo> getAllUserLinks()
        {
            return _listItemInfos;
        }

        public string GetShortcutTarget(string lnkFilePath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkFilePath);
            return shortcut.TargetPath;
        }
        public BitmapImage GetIconImageFromExecutable(string exePath, int iconIndex)
        {
            ExtractIconEx(exePath, iconIndex, out IntPtr largeIcon, out _, 1);

            Icon icon = Icon.FromHandle(largeIcon);
            BitmapImage bitmapImage;
            using (Bitmap bitmap = icon.ToBitmap())
            {
                using (MemoryStream memory = new MemoryStream())
                {
                    bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                    memory.Position = 0;
                    bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memory;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();
                }
            }
            DestroyIcon(largeIcon);
            return bitmapImage;
        }
    }
}
