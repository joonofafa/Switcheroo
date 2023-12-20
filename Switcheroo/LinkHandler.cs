using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Media.Imaging;
using System.Windows.Documents;
using System.Diagnostics;
using System.Windows;
using System.Windows.Forms;
using IWshRuntimeLibrary;

namespace Switcheroo {
    class LinkHandler {

        private List<ListItemInfo> _listItemInfos = new List<ListItemInfo>();

        public void cacheLinks()
        {
            string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            if (Directory.Exists(startMenuPath))
            {
                List<string> lnkFiles = FindLnkFiles(startMenuPath);
                foreach (var file in lnkFiles)
                {
                    Trace.WriteLine("file = [" + file + "]");
                    ListItemInfo listItemInfo = new ListItemInfo
                    {
                        FormattedTitle = Path.GetFileNameWithoutExtension(file),
                        FormattedSubTitle = Path.GetFileNameWithoutExtension(GetShortcutTarget(file)),
                        ImageSource = GetIconImageFromShortcut(file)
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
            // 바로 가기에서 아이콘 위치 추출
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
            string iconLocation = shortcut.IconLocation;

            // 아이콘 위치가 비어있지 않은지 확인
            if (string.IsNullOrEmpty(iconLocation))
            {
                // 기본 아이콘 처리 또는 오류 처리
                return null;
            }

            // 아이콘 위치에서 파일 경로와 인덱스 분리
            string[] iconPathParts = iconLocation.Split(',');
            string iconPath = iconPathParts[0];
            int iconIndex = (iconPathParts.Length > 1) ? int.Parse(iconPathParts[1]) : 0;

            // 아이콘 추출
            Icon icon;
            try
            {
                icon = Icon.ExtractAssociatedIcon(iconPath);
                Bitmap bitmap = icon.ToBitmap();
                using (MemoryStream memory = new MemoryStream())
                {
                    bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                    memory.Position = 0;
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memory;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();
                    Trace.WriteLine("* 이미지 등록 성공! [" + shortcutPath + "]");
                    return bitmapImage;
                }
            }
            catch (ArgumentException)
            {
                Trace.WriteLine("*! 이미지 등록 실패(경로) [" + shortcutPath + "]");
                return null;
            }
            catch (FileNotFoundException)
            {
                Trace.WriteLine("*! 이미지 등록 실패(파일 없음) [" + shortcutPath + "]");
                return null;
            }
        }

        public List<ListItemInfo> getAllUserLinks()
        {
            return _listItemInfos;
        }

        static string GetShortcutTarget(string lnkFilePath)
        {
            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(lnkFilePath);
            return shortcut.TargetPath;
        }
    }
}
