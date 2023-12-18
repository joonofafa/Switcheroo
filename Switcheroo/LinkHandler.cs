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
using ShellLink;
using System.Windows;
using System.Windows.Forms;

namespace Switcheroo {
    class LinkHandler {

        private List<ListItemInfo> _listItemInfos = new List<ListItemInfo>();

        public void cacheLinks()
        {
            string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            if (Directory.Exists(startMenuPath))
            {
                string[] shortcuts = Directory.GetFiles(startMenuPath, "*.lnk");
                foreach (var shortcut in shortcuts)
                {
                    //ShellLink.Shortcut shortItem = ShellLink.Shortcut.ReadFromFile(shortcut);
                    string fileName = Path.GetFileName(shortcut);
                    ListItemInfo listItemInfo = new ListItemInfo
                    {
                        FormattedTitle = fileName.Replace(".lnk", ""),
                        FormattedSubTitle = fileName.Replace(".lnk", ""),
                        //FormattedSubTitle = shortItem.StringData.IconLocation
                    };
                    //System.Windows.MessageBox.Show(shortcut + "," + listItemInfo.FormattedTitle);
                    _listItemInfos.Add(listItemInfo);
                    /*
                    ListItemInfo listItemInfo = new ListItemInfo
                    {
                        FormattedTitle = GetLnkTarget(shortcut),
                        FormattedSubTitle = GetLnkTarget(shortcut)
                    };
                    listItemInfo.ImageSource = ConvertIconToBitmapImage(GetLnkIcon(shortcut).SmallIcon);

                    Debug.WriteLine("ITEM : " + listItemInfo.FormattedTitle);
                    _allUserLinks.Add(listItemInfo);*/
                }
            }
            
        }

        public List<ListItemInfo> getAllUserLinks()
        {
            return _listItemInfos;
        }
    }
}
