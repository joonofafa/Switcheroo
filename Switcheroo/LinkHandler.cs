using IniParser;
using IniParser.Model;
using IWshRuntimeLibrary;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Switcheroo {

    public class ExternalData {
        public string Name { get; set; }
        public string Key { get; set; }
        public string Uri { get; set; }
    }

    class LinkHandler {

        private readonly ConcurrentBag<ListItemInfo> _listExecInfo = new ConcurrentBag<ListItemInfo>();
        private List<ListItemInfo> _listSearchInfo = new List<ListItemInfo>();
        private static readonly string ICON_DEFAULT_PATH = "C:\\Windows\\System32\\shell32.dll";
        private static readonly int ICON_DEFAULT_INDEX = 2;
        private static readonly int ICON_WEB_INDEX = 13;

        public async Task CacheExecutableLinksList()
        {
            // Clear existing data
            while (!_listExecInfo.IsEmpty)
            {
                _listExecInfo.TryTake(out _);
            }

            // Collect all file paths from both locations
            var allLnkFiles = new List<string>();
            var allUrlFiles = new List<string>();

            string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
            string programsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs");

            // Collect files from both locations
            await Task.Run(() =>
            {
                if (Directory.Exists(startMenuPath))
                {
                    allLnkFiles.AddRange(FindLnkFiles(startMenuPath));
                    allUrlFiles.AddRange(FindUrlFiles(startMenuPath));
                }

                if (Directory.Exists(programsPath))
                {
                    allLnkFiles.AddRange(FindLnkFiles(programsPath));
                    allUrlFiles.AddRange(FindUrlFiles(programsPath));
                }
            });

            // Process .lnk files in parallel
            await Task.Run(() =>
            {
                Parallel.ForEach(allLnkFiles, file =>
                {
                    try
                    {
                        var itemInfo = GetLinkItemInfoFromShortcut(file);
                        if (itemInfo != null)
                        {
                            _listExecInfo.Add(itemInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log exception if needed, but continue processing other files
                        Console.WriteLine($"Error processing .lnk file {file}: {ex.Message}");
                    }
                });
            });

            // Process .url files in parallel
            await Task.Run(() =>
            {
                Parallel.ForEach(allUrlFiles, file =>
                {
                    try
                    {
                        var itemInfo = GetUrlFromShortcut(file);
                        if (itemInfo != null)
                        {
                            _listExecInfo.Add(itemInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log exception if needed, but continue processing other files
                        Console.WriteLine($"Error processing .url file {file}: {ex.Message}");
                    }
                });
            });

            // These can also run in parallel
            await Task.Run(() =>
            {
                Parallel.Invoke(
                    () => MakeUWPAppList(),
                    () => MakeSearchList()
                );
            });
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
                    var listItem = new ListItemInfo
                    {
                        FormattedTitle = app.Name,
                        FormattedSubTitle = app.Key,
                        TagData = app.Uri,
                        IsUrl = false,
                        IconPath = ICON_DEFAULT_PATH,
                        IconIndex = ICON_DEFAULT_INDEX,
                        IsDefaultIcon = true
                    };
                    _listExecInfo.Add(listItem);
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
                    var listItem = new ListItemInfo
                    {
                        FormattedTitle = app.Name,
                        FormattedSubTitle = app.Key,
                        TagData = app.Uri,
                        IsUrl = true,
                        IconPath = ICON_DEFAULT_PATH,
                        IconIndex = ICON_WEB_INDEX,
                        IsDefaultIcon = true
                    };
                    _listSearchInfo.Add(listItem);
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
            try
            {
                var parser = new FileIniDataParser();
                IniData data = parser.ReadFile(filePath);
                
                if (data == null || !data.Sections.ContainsSection("InternetShortcut"))
                {
                    return null;
                }

                var internetShortcutSection = data["InternetShortcut"];
                
                // URL is required
                if (!internetShortcutSection.ContainsKey("URL") || string.IsNullOrEmpty(internetShortcutSection["URL"]))
                {
                    return null;
                }

                string url = internetShortcutSection["URL"].Trim();
                string iconPath = "";

                // IconFile is optional
                if (internetShortcutSection.ContainsKey("IconFile") && !string.IsNullOrEmpty(internetShortcutSection["IconFile"]))
                {
                    iconPath = internetShortcutSection["IconFile"].Trim();
                }

                ListItemInfo listItemInfo = new ListItemInfo
                {
                    FormattedTitle = Path.GetFileNameWithoutExtension(filePath),
                    FormattedSubTitle = Path.GetFileNameWithoutExtension(filePath),
                    TagData = url,
                    IsUrl = true
                };

                if (!string.IsNullOrEmpty(iconPath))
                {
                    listItemInfo.IconPath = iconPath;
                    listItemInfo.IconIndex = 0;
                    listItemInfo.IsDefaultIcon = false;
                }
                else
                {
                    listItemInfo.IconPath = ICON_DEFAULT_PATH;
                    listItemInfo.IconIndex = ICON_DEFAULT_INDEX;
                    listItemInfo.IsDefaultIcon = true;
                }

                return listItemInfo;
            }
            catch (Exception ex)
            {
                // Log the specific error for debugging
                Console.WriteLine($"Error processing URL file {filePath}: {ex.Message}");
                return null;
            }
        }

        private ListItemInfo GetLinkItemInfoFromShortcut(string shortcutPath)
        {
            WshShell shell = null;
            IWshShortcut shortcut = null;

            try
            {
                shell = new WshShell();
                shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                
                if (shortcut == null)
                {
                    return null;
                }

                string iconLocation = shortcut.IconLocation ?? "";
                string targetPath = shortcut.TargetPath ?? "";
                
                ListItemInfo listItemInfo = new ListItemInfo
                {
                    FormattedTitle = Path.GetFileNameWithoutExtension(shortcutPath),
                    FormattedSubTitle = Path.GetFileNameWithoutExtension(targetPath),
                    TagData = shortcutPath,
                    ImageSource = null,
                    IsUrl = false
                };

                if (string.IsNullOrEmpty(iconLocation))
                {
                    listItemInfo.IconPath = ICON_DEFAULT_PATH;
                    listItemInfo.IconIndex = ICON_DEFAULT_INDEX;
                    listItemInfo.IsDefaultIcon = true;
                    return listItemInfo;
                }

                try
                {
                    string[] iconPathParts = iconLocation.Split(',');
                    string iconPath = (!string.IsNullOrEmpty(iconPathParts[0].Trim()) && iconPathParts[0].Trim().Length > 1) 
                                    ? iconPathParts[0].Trim() 
                                    : targetPath;
                    
                    int iconIndex = 0;
                    if (iconPathParts.Length > 1 && int.TryParse(iconPathParts[1].Trim(), out int parsedIndex))
                    {
                        iconIndex = parsedIndex;
                    }

                    listItemInfo.IconPath = iconPath;
                    listItemInfo.IconIndex = iconIndex;
                    listItemInfo.IsDefaultIcon = false;
                }
                catch (Exception)
                {
                    listItemInfo.IconPath = ICON_DEFAULT_PATH;
                    listItemInfo.IconIndex = ICON_DEFAULT_INDEX;
                    listItemInfo.IsDefaultIcon = true;
                }

                return listItemInfo;
            }
            catch (Exception ex)
            {
                // Log the specific error for debugging
                Console.WriteLine($"Error processing shortcut file {shortcutPath}: {ex.Message}");
                return null;
            }
            finally
            {
                // COM 객체 리소스 해제
                if (shortcut != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
                }
                if (shell != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
                }
            }
        }

        public List<ListItemInfo> GetAllExecuteableLinksList()
        {
            return _listExecInfo.ToList();
        }

        public List<ListItemInfo> GetAllSearchList()
        {
            return _listSearchInfo;
        }

    }
}
