
using IniParser;
using IniParser.Model;
using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
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

        // Loading state and cancellation support
        private CancellationTokenSource _cancellationTokenSource;
        private readonly object _lockObject = new object();
        public bool IsLoading { get; private set; }
        public bool IsLoaded { get; private set; }

        /// <summary>
        /// Cancels any ongoing caching operation
        /// </summary>
        public void CancelLoading()
        {
            lock (_lockObject)
            {
                if (_cancellationTokenSource != null && !_cancellationTokenSource.IsCancellationRequested)
                {
                    _cancellationTokenSource.Cancel();
                }
            }
        }

        /// <summary>
        /// Reload search_list.json and uwp_list.json configurations
        /// </summary>
        public void ReloadSearchAndUwpConfigs()
        {
            // Clear and reload search list
            _listSearchInfo.Clear();
            MakeSearchList();
        }

        public async Task CacheExecutableLinksList(CancellationToken externalToken = default)
        {
            // Prevent multiple simultaneous loading operations
            lock (_lockObject)
            {
                if (IsLoading)
                {
                    return;
                }
                IsLoading = true;
                _cancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(externalToken);
            }

            var token = _cancellationTokenSource.Token;

            try
            {
                // Clear existing data
                while (!_listExecInfo.IsEmpty)
                {
                    _listExecInfo.TryTake(out _);
                }
                _listSearchInfo.Clear();

                // Collect all file paths from both locations
                var allLnkFiles = new List<string>();
                var allUrlFiles = new List<string>();

                string startMenuPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
                string programsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs");

                // Collect files from both locations
                await Task.Run(() =>
                {
                    token.ThrowIfCancellationRequested();

                    if (Directory.Exists(startMenuPath))
                    {
                        allLnkFiles.AddRange(FindLnkFiles(startMenuPath));
                        allUrlFiles.AddRange(FindUrlFiles(startMenuPath));
                    }

                    token.ThrowIfCancellationRequested();

                    if (Directory.Exists(programsPath))
                    {
                        allLnkFiles.AddRange(FindLnkFiles(programsPath));
                        allUrlFiles.AddRange(FindUrlFiles(programsPath));
                    }
                }, token);

                token.ThrowIfCancellationRequested();

                // Process .lnk files - use STA thread for COM objects
                // WshShell is a STA COM object, so we process sequentially to avoid threading issues
                await Task.Run(() =>
                {
                    foreach (var file in allLnkFiles)
                    {
                        if (token.IsCancellationRequested)
                            break;

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
                            System.Diagnostics.Debug.WriteLine($"Error processing .lnk file {file}: {ex.Message}");
                        }
                    }
                }, token);

                token.ThrowIfCancellationRequested();

                // Process .url files in parallel (no COM objects involved)
                await Task.Run(() =>
                {
                    var options = new ParallelOptions 
                    { 
                        CancellationToken = token,
                        MaxDegreeOfParallelism = Environment.ProcessorCount 
                    };

                    try
                    {
                        Parallel.ForEach(allUrlFiles, options, file =>
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
                                System.Diagnostics.Debug.WriteLine($"Error processing .url file {file}: {ex.Message}");
                            }
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        // Expected when cancelled
                    }
                }, token);

                token.ThrowIfCancellationRequested();

                // Load UWP and Search lists
                await Task.Run(() =>
                {
                    MakeUWPAppList();
                    token.ThrowIfCancellationRequested();
                    MakeSearchList();
                }, token);

                IsLoaded = true;
            }
            catch (OperationCanceledException)
            {
                // Loading was cancelled, this is expected behavior
                System.Diagnostics.Debug.WriteLine("Link caching was cancelled.");
                IsLoaded = false;
            }
            catch (Exception ex)
            {
                // Unexpected error during loading
                System.Diagnostics.Debug.WriteLine($"Error during link caching: {ex.Message}");
                IsLoaded = false;
            }
            finally
            {
                lock (_lockObject)
                {
                    IsLoading = false;
                    _cancellationTokenSource?.Dispose();
                    _cancellationTokenSource = null;
                }
            }
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
            try
            {
                string targetPath = "";
                string iconLocation = "";
                int iconIndex = 0;

                // Use Shell32 IShellLink interface instead of WshShell COM object
                // This is thread-safe and doesn't require STA
                ShellLink shellLink = new ShellLink();
                try
                {
                    ((IPersistFile)shellLink).Load(shortcutPath, 0);
                    
                    StringBuilder targetBuilder = new StringBuilder(260);
                    ((IShellLinkW)shellLink).GetPath(targetBuilder, targetBuilder.Capacity, IntPtr.Zero, SLGP_FLAGS.SLGP_RAWPATH);
                    targetPath = targetBuilder.ToString();

                    StringBuilder iconBuilder = new StringBuilder(260);
                    ((IShellLinkW)shellLink).GetIconLocation(iconBuilder, iconBuilder.Capacity, out iconIndex);
                    iconLocation = iconBuilder.ToString();
                    
                    // Expand environment variables in paths (e.g., %SystemRoot%)
                    if (!string.IsNullOrEmpty(iconLocation))
                    {
                        iconLocation = Environment.ExpandEnvironmentVariables(iconLocation);
                    }
                    if (!string.IsNullOrEmpty(targetPath))
                    {
                        targetPath = Environment.ExpandEnvironmentVariables(targetPath);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(shellLink);
                }

                ListItemInfo listItemInfo = new ListItemInfo
                {
                    FormattedTitle = Path.GetFileNameWithoutExtension(shortcutPath),
                    FormattedSubTitle = !string.IsNullOrEmpty(targetPath) ? Path.GetFileNameWithoutExtension(targetPath) : "",
                    TagData = shortcutPath,
                    ImageSource = null,
                    IsUrl = false
                };

                if (!string.IsNullOrEmpty(iconLocation))
                {
                    listItemInfo.IconPath = iconLocation;
                    listItemInfo.IconIndex = iconIndex;
                    listItemInfo.IsDefaultIcon = false;
                }
                else if (!string.IsNullOrEmpty(targetPath))
                {
                    listItemInfo.IconPath = targetPath;
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
                System.Diagnostics.Debug.WriteLine($"Error processing shortcut file {shortcutPath}: {ex.Message}");
                return null;
            }
        }

        #region Shell32 IShellLink Interop

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink { }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLinkW
        {
            // Note: IUnknown methods are automatically handled by InterfaceIsIUnknown
            // Do NOT include QueryInterface, AddRef, Release here!
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, SLGP_FLAGS fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out ushort pwHotkey);
            void SetHotkey(ushort wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
            void Resolve(IntPtr hwnd, uint fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [Flags]
        private enum SLGP_FLAGS
        {
            SLGP_SHORTPATH = 0x1,
            SLGP_UNCPRIORITY = 0x2,
            SLGP_RAWPATH = 0x4
        }

        #endregion

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
