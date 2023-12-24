using IniParser.Model;
using IniParser;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Media.Imaging;

namespace Switcheroo {

    public class ExternalData {
        public string Name { get; set; }
        public string Key { get; set; }
        public string Uri { get; set; }
    }

    class LinkHandler {

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        extern static int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr largeIcon, out IntPtr smallIcon, int nIcons);
        [DllImport("user32.dll", SetLastError = true)]
        extern static bool DestroyIcon(IntPtr hIcon);

        private List<ListItemInfo> _listExecInfo = new List<ListItemInfo>();
        private List<ListItemInfo> _listSearchInfo = new List<ListItemInfo>();
        private static readonly string ICON_DEFAULT_PATH = "C:\\Windows\\System32\\shell32.dll";
        private static readonly int ICON_DEFAULT_INDEX = 2;
        private static readonly int ICON_WEB_INDEX = 13;

        public void CacheExecutableLinksList()
        {
            string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            List<string> lnkFiles;
            ListItemInfo itemInfo;

            if (Directory.Exists(startMenuPath))
            {
                lnkFiles = FindLnkFiles(startMenuPath);
                foreach (var file in lnkFiles)
                {
                    itemInfo = GetLinkItemInfoFromShortcut(file);
                    if (itemInfo != null)
                    {
                        _listExecInfo.Add(itemInfo);
                    }
                }
                
                lnkFiles = FindUrlFiles(startMenuPath);
                foreach (var file in lnkFiles)
                {
                    itemInfo = GetUrlFromShortcut(file);
                    if (itemInfo != null)
                    {
                        _listExecInfo.Add(itemInfo);
                    }
                }
            }

            string programsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs");
            if (Directory.Exists(programsPath))
            {
                lnkFiles = FindLnkFiles(programsPath);
                foreach (var file in lnkFiles)
                {
                    itemInfo = GetLinkItemInfoFromShortcut(file);
                    if (itemInfo != null)
                    {
                        _listExecInfo.Add(itemInfo);
                    }
                }
                
                lnkFiles = FindUrlFiles(programsPath);
                foreach (var file in lnkFiles)
                {
                    itemInfo = GetUrlFromShortcut(file);
                    if (itemInfo != null)
                    {
                        _listExecInfo.Add(itemInfo);
                    }
                }
            }

            MakeUWPAppList();
            MakeSearchList();
        }

        private void MakeUWPAppList()
        {
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "uwp_list.json");

            try
            {
                string json = System.IO.File.ReadAllText(jsonFilePath);
                List<ExternalData> uwpApps = JsonConvert.DeserializeObject<List<ExternalData>>(json);

                foreach (var app in uwpApps)
                {
                    _listExecInfo.Add(new ListItemInfo(app.Name, app.Key, GetIconImageFromExecutable(ICON_DEFAULT_PATH, ICON_DEFAULT_INDEX), app.Uri, false));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception : {ex.Message}");
            }
        }

        private void MakeSearchList()
        {
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "search_list.json");

            try
            {
                string json = System.IO.File.ReadAllText(jsonFilePath);
                List<ExternalData> searhSites = JsonConvert.DeserializeObject<List<ExternalData>>(json);

                foreach (var app in searhSites)
                {
                    _listSearchInfo.Add(new ListItemInfo(app.Name, app.Key, GetIconImageFromExecutable(ICON_DEFAULT_PATH, ICON_WEB_INDEX), app.Uri, true));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception : {ex.Message}");
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

        private List<string> FindUrlFiles(string directoryPath)
        {
            DirectoryInfo di = new DirectoryInfo(directoryPath);
            return di.EnumerateFiles("*.url", SearchOption.AllDirectories)
                     .Where(fi => !fi.Name.Contains("Uninstall"))
                     .Select(fi => fi.FullName)
                     .ToList();
        }

        private ListItemInfo GetUrlFromShortcut(string filePath)
        {
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile(filePath);
            ListItemInfo listItemInfo = new ListItemInfo();

            try
            {
                string url = data["InternetShortcut"]["URL"].Trim();
                string iconPath = data["InternetShortcut"]["IconFile"].Trim();

                listItemInfo.FormattedTitle = Path.GetFileNameWithoutExtension(filePath);
                listItemInfo.FormattedSubTitle = Path.GetFileNameWithoutExtension(filePath);
                listItemInfo.TagData = url;
                if (iconPath.Length > 0)
                {
                    listItemInfo.ImageSource = GetIconImageFromIconFile(iconPath);
                }
                else
                {
                    listItemInfo.ImageSource = GetIconImageFromExecutable(ICON_DEFAULT_PATH, ICON_DEFAULT_INDEX);
                }
                listItemInfo.IsUrl = true;
            }
            catch
            {
                listItemInfo = null;
            }
            return listItemInfo;
        }

        private ListItemInfo GetLinkItemInfoFromShortcut(string shortcutPath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            string iconLocation = shortcut.IconLocation;
            string targetPath = shortcut.TargetPath;
            ListItemInfo listItemInfo = new ListItemInfo();

            listItemInfo.FormattedTitle = Path.GetFileNameWithoutExtension(shortcutPath);
            listItemInfo.FormattedSubTitle = Path.GetFileNameWithoutExtension(targetPath);
            listItemInfo.TagData = shortcutPath;
            listItemInfo.ImageSource = null;
            listItemInfo.IsUrl = false;

            if (string.IsNullOrEmpty(iconLocation))
            {
                return listItemInfo;
            }

            string[] iconPathParts = iconLocation.Split(',');
            string iconPath = (iconPathParts[0].Trim().Length > 1) ? iconPathParts[0].Trim() : targetPath;
            int iconIndex = (iconPathParts.Length > 1) ? int.Parse(iconPathParts[1]) : 0;

            try
            {
                listItemInfo.ImageSource = GetIconImageFromExecutable(iconPath, iconIndex);
            }
            catch (Exception)
            {
                listItemInfo.ImageSource = GetIconImageFromExecutable(ICON_DEFAULT_PATH, ICON_DEFAULT_INDEX);
            }
            return listItemInfo;
        }

        public List<ListItemInfo> GetAllExecuteableLinksList()
        {
            return _listExecInfo;
        }

        public List<ListItemInfo> GetAllSearchList()
        {
            return _listSearchInfo;
        }

        public BitmapImage GetIconImageFromExecutable(string exePath, int iconIndex)
        {
            ExtractIconEx(exePath, iconIndex, out IntPtr largeIcon, out _, 1);

            Icon icon = Icon.FromHandle(largeIcon);
            BitmapImage bitmapImage = null;

            try
            {
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
            }
            catch (Exception) { }
            return bitmapImage;
        }

        public BitmapImage GetIconImageFromIconFile(string iconPath)
        {
            BitmapImage bitmapImage = null;

            try
            {
                using (Icon ico = new Icon(iconPath))
                {
                    using (Bitmap bmp = ico.ToBitmap())
                    {
                        var memoryStream = new MemoryStream();
                        bmp.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                        memoryStream.Position = 0;

                        bitmapImage = new BitmapImage();
                        bitmapImage.BeginInit();
                        bitmapImage.StreamSource = memoryStream;
                        bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                        bitmapImage.EndInit();
                        bitmapImage.Freeze();
                    }
                }
            }
            catch (Exception) { }
            return bitmapImage;
        }
    }
}
