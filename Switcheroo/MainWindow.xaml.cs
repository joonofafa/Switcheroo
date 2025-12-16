/*
 * Switcheroo - The incremental-search task switcher for Windows.
 * http://www.switcheroo.io/
 * Copyright 2009, 2010 James Sulak
 * Copyright 2014 Regin Larsen
 * 
 * Switcheroo is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * Switcheroo is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with Switcheroo.  If not, see <http://www.gnu.org/licenses/>.
 */

using ManagedWinapi;
using ManagedWinapi.Windows;
using Newtonsoft.Json;
using Switcheroo.Core;
using Switcheroo.Core.Matchers;
using Switcheroo.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Application = System.Windows.Application;
using MenuItem = System.Windows.Forms.MenuItem;
using MessageBox = System.Windows.MessageBox;

namespace Switcheroo {
    public class IgnoreItem {
        public string Name { get; set; }
    }

    /// <summary>
    /// Enum for macro functions
    /// </summary>
    public enum MacroFunction
    {
        PATH_MACRO,     // Navigate to specified path
        // Future functions can be added here
    }

    /// <summary>
    /// Execute action for function key macro
    /// </summary>
    public class ExecuteAction
    {
        [JsonProperty("func")]
        public string Func { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }
    }

    /// <summary>
    /// Function key macro definition
    /// </summary>
    public class FunctionKeyMacro
    {
        [JsonProperty("key")]
        public string Key { get; set; }

        [JsonProperty("process_list")]
        public string ProcessList { get; set; }

        [JsonProperty("execute_list")]
        public ExecuteAction ExecuteList { get; set; }
    }

    public partial class MainWindow : Window {
        private WindowCloser _windowCloser;                                             // Used to close windows
        private List<AppWindowViewModel> _unfilteredWindowList;                         // All windows
        private ObservableCollection<AppWindowViewModel> _filteredWindowList;           // Filtered windows
        private List<ListItemInfo> _unfilteredLinkList;                                 // All links
        private List<ListItemInfo> _filteredLinkList;                                   // Filtered links
        private List<ListItemInfo> _unfilteredWebList;                                  // All links
        private List<ListItemInfo> _filteredWebList;                                    // Filtered links
        private List<IgnoreItem> _ignoreList;                                           // Ignore List Items
        private NotifyIcon _notifyIcon;                                                 // System tray icon
        private HotKey _hotkey;                                                         // Hotkey for activating Switcheroo
        private HotKeyForExecuter _hotkeyForExecuter;                                   // Hotkey for activating Switcheroo
        private LinkHandler _linkHandler = null;                                        // Link Handler
        private bool _isLinkQuiryMode = false;                                          // Link Mode
        private bool _isQuirySearchMode = false;                                        // Quiry Search Mode
        private bool _isFileExplorerMode = false;                                       // File Explorer Mode
        private bool _suppressTextChanged = false;                                      // Suppress TextChanged event
        private string _currentExplorerPath = "";                                       // Current explorer path
        private List<FunctionKeyMacro> _functionKeyMacros;                              // Function key macros

        public static readonly RoutedUICommand CloseWindowCommand = new RoutedUICommand();
        public static readonly RoutedUICommand SwitchToWindowCommand = new RoutedUICommand();
        public static readonly RoutedUICommand ScrollListDownCommand = new RoutedUICommand();
        public static readonly RoutedUICommand ScrollListUpCommand = new RoutedUICommand();
        private OptionsWindow _optionsWindow;
        private AboutWindow _aboutWindow;
        private AltTabHook _altTabHook;
        private SystemWindow _foregroundWindow;
        private bool _altTabAutoSwitch;

        public MainWindow()
        {
            InitializeComponent();
            SetUpKeyBindings();
            SetUpNotifyIcon();
            SetUpHotKey();
            SetUpAltTabHook();
            LoadIgnoreItems();
            LoadFunctionKeyMacros();
            Opacity = 0;
        }

        /// =================================

        #region Private Methods

        /// =================================

        private void SetUpKeyBindings()
        {
            // Enter and Esc bindings are not executed before the keys have been released.
            // This is done to prevent that the window being focused after the key presses
            // to get 'KeyUp' messages.

            KeyDown += (sender, args) =>
            {
                // Opacity is set to 0 right away so it appears that action has been taken right away...
                if (args.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
                {
                    Opacity = 0;
                }
                else if (args.Key == Key.Escape)
                {
                    Opacity = 0;
                }
                else if (args.SystemKey == Key.S && Keyboard.Modifiers.HasFlag(ModifierKeys.Alt))
                {
                    _altTabAutoSwitch = false;
                    tb.Text = "";
                    tb.IsEnabled = true;
                    tb.Focus();
                }
            };

            // PreviewKeyDown to intercept Tab key and Function keys
            PreviewKeyDown += (sender, args) =>
            {
                // Tab key - Bash-style auto-completion or directory navigation
                if (args.Key == Key.Tab && _isFileExplorerMode && !Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                {
                    HandleTabKeyInFileExplorer();
                    args.Handled = true;
                }
                // Function keys (F1-F12) - execute macros in Executer mode
                else if (_isLinkQuiryMode && args.Key >= Key.F1 && args.Key <= Key.F12)
                {
                    if (TryExecuteFunctionKeyMacro(args.Key))
                    {
                        args.Handled = true;
                    }
                }
            };

            KeyUp += (sender, args) =>
            {
                // ... But only when the keys are release, the action is actually executed
                if (args.Key == Key.Enter && !Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
                {
                    if (_isLinkQuiryMode)
                    {
                        Execute();
                    }
                    else
                    {
                        Switch();
                    }
                }
                else if (args.Key == Key.Escape)
                {
                    HideWindow();
                }
                else if (args.SystemKey == Key.LeftAlt && !Keyboard.Modifiers.HasFlag(ModifierKeys.Control))
                {
                    if (_isLinkQuiryMode)
                    {
                        Execute();
                    }
                    else
                    {
                        Switch();
                    }
                }
                else if (args.Key == Key.LeftAlt && _altTabAutoSwitch)
                {
                    if (_isLinkQuiryMode)
                    {
                        Execute();
                    }
                    else
                    {
                        Switch();
                    }
                }
            };
        }

        private void SetUpHotKey()
        {
            _hotkey = new HotKey();
            _hotkey.LoadSettings();

            _hotkeyForExecuter = new HotKeyForExecuter();
            _hotkeyForExecuter.LoadSettings();

            Application.Current.Properties["hotkey"] = _hotkey;
            Application.Current.Properties["lnkHotkey"] = _hotkeyForExecuter;

            _hotkey.HotkeyPressed += Hotkey_HotkeyPressed;
            _hotkeyForExecuter.HotkeyPressed += Hotkey_HotkeyForExecuterPressed;
            try
            {
                _hotkey.Enabled = Settings.Default.EnableHotKey;
                _hotkeyForExecuter.Enabled = Settings.Default.EnableLnkHotkey;
            }
            catch (HotkeyAlreadyInUseException)
            {
                var boxText = "The current hotkey for activating Switcheroo is in use by another program." +
                              Environment.NewLine +
                              Environment.NewLine +
                              "You can change the hotkey by right-clicking the Switcheroo icon in the system tray and choosing 'Options'.";
                MessageBox.Show(boxText, "Hotkey already in use", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SetUpAltTabHook()
        {
            _altTabHook = new AltTabHook();
            _altTabHook.Pressed += AltTabPressed;
        }

        private void SetUpNotifyIcon()
        {
            var icon = Properties.Resources.icon;

            var runOnStartupMenuItem = new MenuItem("Run on Startup", (s, e) => RunOnStartup(s as MenuItem))
            {
                Checked = new AutoStart().IsEnabled
            };

            _notifyIcon = new NotifyIcon
            {
                Text = "Switcheroo",
                Icon = icon,
                Visible = true,
                ContextMenu = new System.Windows.Forms.ContextMenu(new[]
                {
                    new MenuItem("Options", (s, e) => Options()),
                    runOnStartupMenuItem,
                    new MenuItem("About", (s, e) => About()),
                    new MenuItem("Exit", (s, e) => Quit())
                })
            };
        }

        private static void RunOnStartup(MenuItem menuItem)
        {
            try
            {
                var autoStart = new AutoStart
                {
                    IsEnabled = !menuItem.Checked
                };
                menuItem.Checked = autoStart.IsEnabled;
            }
            catch (AutoStartException e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadIgnoreItems()
        {
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ignore_list.json");

            try
            {
                string json = File.ReadAllText(jsonFilePath);
                _ignoreList = JsonConvert.DeserializeObject<List<IgnoreItem>>(json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception : {ex.Message}");
            }
        }

        /// <summary>
        /// Load function key macros from func_list.json
        /// </summary>
        private void LoadFunctionKeyMacros()
        {
            _functionKeyMacros = new List<FunctionKeyMacro>();
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "func_list.json");

            try
            {
                if (File.Exists(jsonFilePath))
                {
                    string json = File.ReadAllText(jsonFilePath);
                    _functionKeyMacros = JsonConvert.DeserializeObject<List<FunctionKeyMacro>>(json) ?? new List<FunctionKeyMacro>();
                    Debug.WriteLine($"Loaded {_functionKeyMacros.Count} function key macros");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error loading func_list.json: {ex.Message}");
            }
        }

        /// <summary>
        /// Execute function key macro if defined
        /// </summary>
        private bool TryExecuteFunctionKeyMacro(Key key)
        {
            if (_functionKeyMacros == null || _functionKeyMacros.Count == 0)
                return false;

            // Convert Key to string (e.g., Key.F5 -> "F5")
            string keyName = key.ToString();

            // Find matching macro
            var macro = _functionKeyMacros.FirstOrDefault(m => 
                m.Key.Equals(keyName, StringComparison.OrdinalIgnoreCase));

            if (macro == null || macro.ExecuteList == null)
                return false;

            // Execute the macro based on function type
            return ExecuteMacroFunction(macro.ExecuteList);
        }

        /// <summary>
        /// Execute macro function
        /// </summary>
        private bool ExecuteMacroFunction(ExecuteAction action)
        {
            if (action == null || string.IsNullOrEmpty(action.Func))
                return false;

            // Parse function type
            if (Enum.TryParse<MacroFunction>(action.Func, true, out MacroFunction func))
            {
                switch (func)
                {
                    case MacroFunction.PATH_MACRO:
                        return ExecutePathMacro(action.Value);
                    // Future functions can be added here
                    default:
                        return false;
                }
            }

            // Also support the original name with typo (PATH_MARCO)
            if (action.Func.Equals("PATH_MARCO", StringComparison.OrdinalIgnoreCase))
            {
                return ExecutePathMacro(action.Value);
            }

            return false;
        }

        /// <summary>
        /// Execute PATH_MACRO function - navigate to specified path
        /// </summary>
        private bool ExecutePathMacro(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;

            // Ensure we're in Executer mode
            if (!_isLinkQuiryMode)
                return false;

            _suppressTextChanged = true;
            
            // Set the path in text box
            if (Directory.Exists(path))
            {
                tb.Text = path.EndsWith("\\") ? path : path + "\\";
            }
            else if (File.Exists(path))
            {
                tb.Text = path;
            }
            else
            {
                tb.Text = path;
            }
            
            tb.CaretIndex = tb.Text.Length;
            _suppressTextChanged = false;

            // Handle as file explorer input
            HandleFileExplorerInput(tb.Text);

            return true;
        }

        /// <summary>
        /// Populates the window list with the current running windows.
        /// </summary>
        private void LoadData(InitialFocus focus)
        {
            tb.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF68686F"));
            _isLinkQuiryMode = false;
            _unfilteredWindowList = new WindowFinder().GetWindows().Select(window => new AppWindowViewModel(window)).ToList();
            foreach (var ignore in _ignoreList)
            {
                var windowToRemove = _unfilteredWindowList.FirstOrDefault(window => window.WindowTitle.ToUpper().Equals(ignore.Name));
                if (windowToRemove != null)
                {
                    _unfilteredWindowList.Remove(windowToRemove);
                }
            }

            var firstWindow = _unfilteredWindowList.FirstOrDefault();
            var foregroundWindowMovedToBottom = false;

            // Move first window to the bottom of the list if it's related to the foreground window
            if (firstWindow != null && AreWindowsRelated(firstWindow.AppWindow, _foregroundWindow))
            {
                _unfilteredWindowList.RemoveAt(0);
                _unfilteredWindowList.Add(firstWindow);
                foregroundWindowMovedToBottom = true;
            }

            _filteredWindowList = new ObservableCollection<AppWindowViewModel>(_unfilteredWindowList);
            _windowCloser = new WindowCloser();

            foreach (var window in _unfilteredWindowList)
            {
                window.FormattedTitle = new XamlHighlighter().Highlight(new[] { new StringPart(window.AppWindow.Title) });
                window.FormattedSubTitle =
                    new XamlHighlighter().Highlight(new[] { new StringPart(window.AppWindow.ProcessTitle) });
            }

            lb.DataContext = null;
            lb.DataContext = _filteredWindowList;

            FocusItemInList(focus, foregroundWindowMovedToBottom);

            tb.Clear();
            tb.Focus();
            CenterWindow();
            ScrollSelectedItemIntoView();
        }

        private async Task LoadLinkData(InitialFocus focus)
        {
            tb.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF515173"));

            _isLinkQuiryMode = true;
            _isQuirySearchMode = false;

            try
            {
                if (_linkHandler == null)
                {
                    _linkHandler = new LinkHandler();
                }

                // Only cache if not already loaded or currently loading
                if (!_linkHandler.IsLoaded && !_linkHandler.IsLoading)
                {
                    await _linkHandler.CacheExecutableLinksList();
                }
                else if (_linkHandler.IsLoading)
                {
                    // Wait for existing loading to complete with timeout
                    const int maxWaitMs = 10000; // 10 seconds max
                    const int delayMs = 50;
                    int waited = 0;
                    
                    while (_linkHandler.IsLoading && waited < maxWaitMs)
                    {
                        await Task.Delay(delayMs);
                        waited += delayMs;
                    }
                    
                    // If still loading after timeout, cancel and return
                    if (_linkHandler.IsLoading)
                    {
                        _linkHandler.CancelLoading();
                        System.Diagnostics.Debug.WriteLine("LoadLinkData: Timeout waiting for loading to complete");
                        return;
                    }
                }

                // Check if loading was successful
                if (!_linkHandler.IsLoaded)
                {
                    return;
                }

                _unfilteredLinkList = _linkHandler.GetAllExecuteableLinksList();
                _unfilteredWebList = _linkHandler.GetAllSearchList();
                lb.DataContext = null;

                tb.Clear();
                tb.Focus();
                CenterWindow();
                ScrollSelectedItemIntoView();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in LoadLinkData: {ex.Message}");
                // Don't rethrow - just hide window gracefully
            }
        }

        private static bool AreWindowsRelated(SystemWindow window1, SystemWindow window2)
        {
            return window1.HWnd == window2.HWnd || window1.Process.Id == window2.Process.Id;
        }

        private void FocusItemInList(InitialFocus focus, bool foregroundWindowMovedToBottom)
        {
            if (focus == InitialFocus.PreviousItem)
            {
                var previousItemIndex = lb.Items.Count - 1;
                if (foregroundWindowMovedToBottom)
                {
                    previousItemIndex--;
                }

                lb.SelectedIndex = previousItemIndex > 0 ? previousItemIndex : 0;
            }
            else
            {
                lb.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Place the Switcheroo window in the center of the screen
        /// </summary>
        private void CenterWindow()
        {
            // Reset height every time to ensure that resolution changes take effect
            Border.MaxHeight = SystemParameters.PrimaryScreenHeight;

            // Force a rendering before repositioning the window
            // Use SizeToContent.Height to keep width fixed
            SizeToContent = SizeToContent.Manual;
            SizeToContent = SizeToContent.Height;

            // Position the window in the center of the screen
            Left = (SystemParameters.PrimaryScreenWidth / 2) - (ActualWidth / 2);
            Top = (SystemParameters.PrimaryScreenHeight / 2) - (ActualHeight / 2);
        }

        /// <summary>
        /// Switches the window associated with the selected item.
        /// </summary>
        private void Switch()
        {
            foreach (var item in lb.SelectedItems)
            {
                var win = (AppWindowViewModel)item;
                win.AppWindow.SwitchToLastVisibleActivePopup();
            }
            HideWindow();
        }
        private void Execute()
        {
            // Handle file explorer mode
            if (_isFileExplorerMode)
            {
                ExecuteFileExplorerAction();
                return;
            }

            if (!_isQuirySearchMode)
            {
                foreach (var item in lb.SelectedItems)
                {
                    var lnk = (ListItemInfo)item;

                    try
                    {
                        if (lnk.IsUrl)
                        {
                            var psi = new ProcessStartInfo
                            {
                                FileName = lnk.TagData.ToString(),
                                UseShellExecute = true
                            };
                            Process.Start(psi);
                        }
                        else
                        {
                            if (lnk.Argument != null)
                            {
                                var psi = new ProcessStartInfo
                                {
                                    FileName = lnk.TagData.ToString(),
                                    Arguments = lnk.Argument,
                                    UseShellExecute = true
                                };
                                Process.Start(psi);
                            }
                            else
                            {
                                Process.Start(lnk.TagData);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine(ex.Message);
                    }
                }
            }
            else
            {
                if (lb.SelectedItems.Count > 0)
                {
                    var lnk = (ListItemInfo)lb.SelectedItems[0];

                    try
                    {
                        string[] parts = tb.Text.Split(' ');
                        var psi = new ProcessStartInfo
                        {
                            FileName = lnk.TagData.Replace("{word}", String.Join(" ", parts.Skip(1))),
                            UseShellExecute = true
                        };
                        Process.Start(psi);
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine(ex.Message);
                    }
                }
            }
            HideWindow();
        }

        #region File Explorer Methods

        /// <summary>
        /// Get list of available drives
        /// </summary>
        private List<ListItemInfo> GetDriveList()
        {
            var driveList = new List<ListItemInfo>();
            
            foreach (var drive in DriveInfo.GetDrives())
            {
                try
                {
                    if (drive.IsReady)
                    {
                        string driveLabel = string.IsNullOrEmpty(drive.VolumeLabel) 
                            ? $"Local Disk ({drive.Name.TrimEnd('\\')})" 
                            : $"{drive.VolumeLabel} ({drive.Name.TrimEnd('\\')})";
                        
                        driveList.Add(new ListItemInfo
                        {
                            FormattedTitle = driveLabel,
                            FormattedSubTitle = drive.DriveType.ToString(),
                            TagData = drive.Name,
                            IsUrl = false,
                            IconPath = drive.Name,
                            IconIndex = 0,
                            IsDefaultIcon = false
                        });
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error accessing drive: {ex.Message}");
                }
            }
            
            return driveList;
        }

        /// <summary>
        /// Get list of directories and files in the specified path
        /// </summary>
        private List<ListItemInfo> GetDirectoryContents(string path)
        {
            var contentList = new List<ListItemInfo>();
            
            try
            {
                if (!Directory.Exists(path))
                    return contentList;

                // Add parent directory option if not root
                var dirInfo = new DirectoryInfo(path);
                if (dirInfo.Parent != null)
                {
                    contentList.Add(new ListItemInfo
                    {
                        FormattedTitle = "..",
                        FormattedSubTitle = "Parent Directory",
                        TagData = dirInfo.Parent.FullName,
                        IsUrl = false,
                        IconPath = dirInfo.Parent.FullName,
                        IconIndex = 0,
                        IsDefaultIcon = false
                    });
                }

                // Add directories first
                foreach (var dir in Directory.GetDirectories(path))
                {
                    try
                    {
                        var di = new DirectoryInfo(dir);
                        // Skip hidden and system directories
                        if ((di.Attributes & FileAttributes.Hidden) != 0 || 
                            (di.Attributes & FileAttributes.System) != 0)
                            continue;

                        contentList.Add(new ListItemInfo
                        {
                            FormattedTitle = di.Name,
                            FormattedSubTitle = "Directory",
                            TagData = di.FullName,
                            IsUrl = false,
                            IconPath = di.FullName,
                            IconIndex = 0,
                            IsDefaultIcon = false
                        });
                    }
                    catch (UnauthorizedAccessException) { }
                    catch (Exception ex) { Debug.WriteLine($"Error: {ex.Message}"); }
                }

                // Add files
                foreach (var file in Directory.GetFiles(path))
                {
                    try
                    {
                        var fi = new FileInfo(file);
                        // Skip hidden and system files
                        if ((fi.Attributes & FileAttributes.Hidden) != 0 ||
                            (fi.Attributes & FileAttributes.System) != 0)
                            continue;

                        contentList.Add(new ListItemInfo
                        {
                            FormattedTitle = fi.Name,
                            FormattedSubTitle = GetFileSizeString(fi.Length),
                            TagData = fi.FullName,
                            IsUrl = false,
                            IconPath = fi.FullName,
                            IconIndex = 0,
                            IsDefaultIcon = false
                        });
                    }
                    catch (UnauthorizedAccessException) { }
                    catch (Exception ex) { Debug.WriteLine($"Error: {ex.Message}"); }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error getting directory contents: {ex.Message}");
            }

            return contentList;
        }

        /// <summary>
        /// Convert file size to human-readable string
        /// </summary>
        private string GetFileSizeString(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size = size / 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }

        /// <summary>
        /// Check if the input is a valid path pattern (e.g., ":", "C:\", "D:\folder")
        /// </summary>
        private bool IsPathInput(string input, out string normalizedPath)
        {
            normalizedPath = "";
            
            if (string.IsNullOrEmpty(input))
                return false;

            // Check for ":" only - show drives
            if (input == ":")
            {
                normalizedPath = ":";
                return true;
            }

            // Check for drive letter patterns like "C:", "C:\", "C:\folder"
            if (input.Length >= 2 && char.IsLetter(input[0]) && input[1] == ':')
            {
                normalizedPath = input;
                if (input.Length == 2)
                    normalizedPath = input + "\\";
                return true;
            }

            return false;
        }

        /// <summary>
        /// Handle file explorer navigation
        /// </summary>
        private void HandleFileExplorerInput(string query)
        {
            _isFileExplorerMode = true;
            _isQuirySearchMode = false;
            
            // Get max items from settings
            int maxItems = Properties.Settings.Default.FileExplorerMaxItems;
            if (maxItems < 10) maxItems = 10;
            if (maxItems > 100) maxItems = 100;

            if (query == ":")
            {
                // Show drive list
                var driveList = GetDriveList();
                lb.DataContext = driveList.Take(maxItems).ToList();
                _currentExplorerPath = "";
            }
            else
            {
                // Show directory contents
                string path = query;
                
                // Normalize path
                if (path.Length == 2 && path[1] == ':')
                    path = path + "\\";

                if (Directory.Exists(path))
                {
                    var contents = GetDirectoryContents(path);
                    
                    // Filter if there's text after the last backslash
                    int lastSlash = query.LastIndexOf('\\');
                    if (lastSlash >= 0 && lastSlash < query.Length - 1)
                    {
                        string filter = query.Substring(lastSlash + 1);
                        string basePath = query.Substring(0, lastSlash + 1);
                        
                        if (Directory.Exists(basePath))
                        {
                            contents = GetDirectoryContents(basePath);
                            contents = contents.Where(item => 
                                item.FormattedTitle.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                            ).Take(maxItems).ToList();
                            _currentExplorerPath = basePath;
                        }
                    }
                    else
                    {
                        _currentExplorerPath = path;
                        contents = contents.Take(maxItems).ToList();
                    }

                    lb.DataContext = contents;
                }
                else
                {
                    // Try to show filtered contents of parent directory
                    int lastSlash = query.LastIndexOf('\\');
                    if (lastSlash >= 0)
                    {
                        string basePath = query.Substring(0, lastSlash + 1);
                        string filter = query.Substring(lastSlash + 1);
                        
                        if (Directory.Exists(basePath))
                        {
                            var contents = GetDirectoryContents(basePath);
                            if (!string.IsNullOrEmpty(filter))
                            {
                                contents = contents.Where(item => 
                                    item.FormattedTitle.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                                ).ToList();
                            }
                            contents = contents.Take(maxItems).ToList();
                            _currentExplorerPath = basePath;
                            lb.DataContext = contents;
                        }
                        else
                        {
                            lb.DataContext = null;
                        }
                    }
                    else
                    {
                        lb.DataContext = null;
                    }
                }
            }

            if (lb.Items.Count > 0)
            {
                lb.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// Execute file explorer action - Enter key opens file or directory in Explorer
        /// </summary>
        private void ExecuteFileExplorerAction()
        {
            if (lb.SelectedItem == null)
                return;

            var item = (ListItemInfo)lb.SelectedItem;
            string path = item.TagData;

            if (Directory.Exists(path))
            {
                // Open directory in Explorer (Enter key)
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = path,
                        UseShellExecute = true
                    });
                    HideWindow();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error opening directory: {ex.Message}");
                }
            }
            else if (File.Exists(path))
            {
                // Open file with shell
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                    HideWindow();
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error opening file: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Navigate into selected directory - Tab key
        /// </summary>
        private void NavigateToSelectedDirectory()
        {
            if (lb.SelectedItem == null)
                return;

            var item = (ListItemInfo)lb.SelectedItem;
            string path = item.TagData;

            if (Directory.Exists(path))
            {
                // Navigate into directory
                _suppressTextChanged = true;
                tb.Text = path.EndsWith("\\") ? path : path + "\\";
                tb.CaretIndex = tb.Text.Length;
                _suppressTextChanged = false;
                
                HandleFileExplorerInput(tb.Text);
            }
        }

        /// <summary>
        /// Handle Tab key in file explorer mode - Bash-style auto-completion
        /// </summary>
        private void HandleTabKeyInFileExplorer()
        {
            string query = tb.Text;
            
            // First, check if the selected item's path matches the current text (arrow key selection case)
            if (lb.SelectedItem != null)
            {
                var selectedItem = lb.SelectedItem as ListItemInfo;
                if (selectedItem != null && selectedItem.FormattedTitle != "..")
                {
                    string selectedPath = selectedItem.TagData;
                    
                    // If text matches selected item's path exactly, navigate into it (like pressing \)
                    if (query.Equals(selectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        if (Directory.Exists(selectedPath))
                        {
                            _suppressTextChanged = true;
                            tb.Text = selectedPath.EndsWith("\\") ? selectedPath : selectedPath + "\\";
                            tb.CaretIndex = tb.Text.Length;
                            _suppressTextChanged = false;
                            
                            HandleFileExplorerInput(tb.Text);
                            return;
                        }
                    }
                }
            }
            
            // Check if there's a filter (partial path after last backslash)
            int lastSlash = query.LastIndexOf('\\');
            bool hasFilter = lastSlash >= 0 && lastSlash < query.Length - 1;
            
            if (hasFilter)
            {
                // Auto-completion mode
                int itemCount = lb.Items.Count;
                
                if (itemCount == 1)
                {
                    // Only one match - auto-complete to this path
                    var item = (ListItemInfo)lb.Items[0];
                    string path = item.TagData;
                    
                    _suppressTextChanged = true;
                    if (Directory.Exists(path))
                    {
                        // It's a directory - append backslash
                        tb.Text = path.EndsWith("\\") ? path : path + "\\";
                    }
                    else
                    {
                        // It's a file - just complete the path
                        tb.Text = path;
                    }
                    tb.CaretIndex = tb.Text.Length;
                    _suppressTextChanged = false;
                    
                    HandleFileExplorerInput(tb.Text);
                }
                else if (itemCount > 1)
                {
                    // Multiple matches - find common prefix and complete to it
                    string commonPrefix = FindCommonPrefix();
                    if (!string.IsNullOrEmpty(commonPrefix))
                    {
                        string basePath = query.Substring(0, lastSlash + 1);
                        string newPath = basePath + commonPrefix;
                        
                        if (newPath != query)
                        {
                            _suppressTextChanged = true;
                            tb.Text = newPath;
                            tb.CaretIndex = tb.Text.Length;
                            _suppressTextChanged = false;
                            
                            HandleFileExplorerInput(tb.Text);
                        }
                        else
                        {
                            // Common prefix equals current query - check if selected item is a directory
                            if (lb.SelectedItem != null)
                            {
                                var selectedItem = lb.SelectedItem as ListItemInfo;
                                if (selectedItem != null && selectedItem.FormattedTitle != ".." && Directory.Exists(selectedItem.TagData))
                                {
                                    _suppressTextChanged = true;
                                    tb.Text = selectedItem.TagData.EndsWith("\\") ? selectedItem.TagData : selectedItem.TagData + "\\";
                                    tb.CaretIndex = tb.Text.Length;
                                    _suppressTextChanged = false;
                                    
                                    HandleFileExplorerInput(tb.Text);
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                // No filter - navigate into selected directory
                NavigateToSelectedDirectory();
            }
        }

        /// <summary>
        /// Find common prefix among all items in the list
        /// </summary>
        private string FindCommonPrefix()
        {
            if (lb.Items.Count == 0)
                return "";

            var items = lb.Items.Cast<ListItemInfo>().ToList();
            
            // Skip ".." entry
            items = items.Where(i => i.FormattedTitle != "..").ToList();
            
            if (items.Count == 0)
                return "";

            string firstTitle = items[0].FormattedTitle;
            int prefixLength = firstTitle.Length;

            foreach (var item in items.Skip(1))
            {
                string title = item.FormattedTitle;
                int minLength = Math.Min(prefixLength, title.Length);
                int matchLength = 0;

                for (int i = 0; i < minLength; i++)
                {
                    if (char.ToLowerInvariant(firstTitle[i]) == char.ToLowerInvariant(title[i]))
                    {
                        matchLength++;
                    }
                    else
                    {
                        break;
                    }
                }

                prefixLength = matchLength;
                if (prefixLength == 0)
                    break;
            }

            return prefixLength > 0 ? firstTitle.Substring(0, prefixLength) : "";
        }

        #endregion

        private void HideWindow()
        {
            if (_windowCloser != null)
            {
                _windowCloser.Dispose();
                _windowCloser = null;
            }

            // Cancel any ongoing link loading operation
            if (_linkHandler != null && _linkHandler.IsLoading)
            {
                _linkHandler.CancelLoading();
            }

            // Reset modes
            _isFileExplorerMode = false;
            _currentExplorerPath = "";

            _altTabAutoSwitch = false;
            Opacity = 0;
            Dispatcher.BeginInvoke(new Action(Hide), DispatcherPriority.Input);
        }

        #endregion

        /// =================================

        #region Right-click menu functions

        /// =================================
        /// <summary>
        /// Show Options dialog.
        /// </summary>
        private void Options()
        {
            if (_optionsWindow == null)
            {
                _optionsWindow = new OptionsWindow
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                _optionsWindow.Closed += (sender, args) => _optionsWindow = null;
                _optionsWindow.ShowDialog();
            }
            else
            {
                _optionsWindow.Activate();
            }
        }

        /// <summary>
        /// Show About dialog.
        /// </summary>
        private void About()
        {
            if (_aboutWindow == null)
            {
                _aboutWindow = new AboutWindow
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen
                };
                _aboutWindow.Closed += (sender, args) => _aboutWindow = null;
                _aboutWindow.ShowDialog();
            }
            else
            {
                _aboutWindow.Activate();
            }
        }

        /// <summary>
        /// Quit Switcheroo
        /// </summary>
        private void Quit()
        {
            _notifyIcon.Dispose();
            _notifyIcon = null;
            _hotkey.Dispose();
            Application.Current.Shutdown();
        }

        #endregion

        /// =================================

        #region Event Handlers

        /// =================================
        private void OnClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            HideWindow();
        }

        private void Hotkey_HotkeyPressed(object sender, EventArgs e)
        {
            if (!Settings.Default.EnableHotKey)
            {
                return;
            }

            if (Visibility != Visibility.Visible)
            {
                tb.IsEnabled = true;

                _foregroundWindow = SystemWindow.ForegroundWindow;
                Show();
                Activate();
                Keyboard.Focus(tb);
                LoadData(InitialFocus.NextItem);
                Opacity = 1;
            }
            else
            {
                HideWindow();
            }
        }

        private async void Hotkey_HotkeyForExecuterPressed(object sender, EventArgs e)
        {
            if (!Settings.Default.EnableLnkHotkey)
            {
                return;
            }

            try
            {
                if (Visibility != Visibility.Visible)
                {
                    // If currently loading, cancel and hide
                    if (_linkHandler != null && _linkHandler.IsLoading)
                    {
                        _linkHandler.CancelLoading();
                        HideWindow();
                        return;
                    }

                    tb.IsEnabled = true;

                    _foregroundWindow = SystemWindow.ForegroundWindow;
                    Show();
                    Activate();
                    Keyboard.Focus(tb);
                    await LoadLinkData(InitialFocus.NextItem);
                    
                    // Only show if loading completed successfully
                    if (_linkHandler != null && _linkHandler.IsLoaded)
                    {
                        Opacity = 1;
                    }
                    else
                    {
                        HideWindow();
                    }
                }
                else
                {
                    HideWindow();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Hotkey_HotkeyForExecuterPressed: {ex.Message}");
                HideWindow();
            }
        }

        private void AltTabPressed(object sender, AltTabHookEventArgs e)
        {
            if (!Settings.Default.AltTabHook)
            {
                // Ignore Alt+Tab presses if the hook is not activated by the user
                return;
            }

            _foregroundWindow = SystemWindow.ForegroundWindow;

            if (_foregroundWindow.ClassName == "MultitaskingViewFrame")
            {
                // If Windows' task switcher is on the screen then don't do anything
                return;
            }

            e.Handled = true;

            if (Visibility != Visibility.Visible)
            {
                tb.IsEnabled = true;

                ActivateAndFocusMainWindow();

                Keyboard.Focus(tb);
                if (e.ShiftDown)
                {
                    LoadData(InitialFocus.PreviousItem);
                }
                else
                {
                    LoadData(InitialFocus.NextItem);
                }

                if (Settings.Default.AutoSwitch && !e.CtrlDown)
                {
                    _altTabAutoSwitch = true;
                    tb.IsEnabled = false;
                    tb.Text = "Press Alt + S to search";
                }

                Opacity = 1;
            }
            else
            {
                if (e.ShiftDown)
                {
                    PreviousItem();
                }
                else
                {
                    NextItem();
                }
            }
        }

        private void ActivateAndFocusMainWindow()
        {
            // What happens below looks a bit weird, but for Switcheroo to get focus when using the Alt+Tab hook,
            // it is needed to simulate an Alt keypress will bring Switcheroo to the foreground. Otherwise Switcheroo
            // will become the foreground window, but the previous window will retain focus, and receive keep getting
            // the keyboard input.
            // http://www.codeproject.com/Tips/76427/How-to-bring-window-to-top-with-SetForegroundWindo

            var thisWindowHandle = new WindowInteropHelper(this).Handle;
            var thisWindow = new AppWindow(thisWindowHandle);

            var altKey = new KeyboardKey(Keys.Alt);
            var altKeyPressed = false;

            // Press the Alt key if it is not already being pressed
            if ((altKey.AsyncState & 0x8000) == 0)
            {
                altKey.Press();
                altKeyPressed = true;
            }

            // Bring the Switcheroo window to the foreground
            Show();
            SystemWindow.ForegroundWindow = thisWindow;
            Activate();

            // Release the Alt key if it was pressed above
            if (altKeyPressed)
            {
                altKey.Release();
            }
        }

        private void TextChanged(object sender, TextChangedEventArgs args)
        {
            if (!tb.IsEnabled || _suppressTextChanged)
            {
                return;
            }

            var query = tb.Text;
            if (!_isLinkQuiryMode)
            {
                var context = new WindowFilterContext<AppWindowViewModel>
                {
                    Windows = _unfilteredWindowList,
                    ForegroundWindowProcessTitle = new AppWindow(_foregroundWindow.HWnd).ProcessTitle
                };
                var filterResults = new WindowFilterer().Filter(context, query).ToList();

                foreach (var filterResult in filterResults)
                {
                    filterResult.AppWindow.FormattedTitle =
                        GetFormattedTitleFromBestResult(filterResult.WindowTitleMatchResults);
                    filterResult.AppWindow.FormattedSubTitle =
                        GetFormattedTitleFromBestResult(filterResult.ProcessTitleMatchResults);
                }

                _filteredWindowList = new ObservableCollection<AppWindowViewModel>(filterResults.Select(r => r.AppWindow));
                lb.DataContext = _filteredWindowList;
                if (lb.Items.Count > 0)
                {
                    lb.SelectedItem = lb.Items[0];
                }
            }
            else
            {
                // Check for file explorer mode (: or path input)
                string normalizedPath;
                if (IsPathInput(query, out normalizedPath))
                {
                    HandleFileExplorerInput(query);
                    return;
                }
                
                // Reset file explorer mode if not a path
                _isFileExplorerMode = false;

                if (_unfilteredLinkList.Count > 0 || _unfilteredWebList.Count > 0)
                {
                    if (query.Length > 0)
                    {
                        if (query.StartsWith("@"))
                        {
                            int spaceIndex = query.IndexOf(' ');
                            if (spaceIndex != -1)
                            {
                                string result = query.Substring(1, spaceIndex).Trim();
                                _filteredWebList = _unfilteredWebList.Where(
                                    item => item.FormattedSubTitle.Equals(result, StringComparison.OrdinalIgnoreCase)).ToList();

                                lb.DataContext = _filteredWebList;
                                if (lb.Items.Count > 0)
                                {
                                    _isQuirySearchMode = true;
                                    lb.SelectedItem = lb.Items[0];
                                }
                                else
                                {
                                    lb.DataContext = null;
                                }
                            }
                            else
                            {
                                lb.DataContext = _unfilteredWebList;
                            }
                        }
                        /*
                        else if (query.StartsWith("?"))
                        {
                            List<ListItemInfo> everythingItem = new List<ListItemInfo>
                            {
                                new ListItemInfo
                                {
                                    FormattedTitle = "Everything Search (" + query.Replace("?", "") + ")",
                                    FormattedSubTitle = "Everything",
                                    TagData = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Everything.exe"),
                                    IsUrl = false,
                                    Argument = "-s " + query.Replace("?", ""),
                                    IconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Everything.exe"),
                                    IconIndex = 0,
                                    IsDefaultIcon = false
                                }
                            };

                            lb.DataContext = everythingItem;
                            lb.SelectedIndex = 0;
                        }*/
                        else
                        {
                            _filteredLinkList = _unfilteredLinkList.Where(
                                    item => item.FormattedTitle.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    item.FormattedSubTitle.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                                    .ToList();   

                            if (_filteredLinkList.Count > 8)
                            {
                                _filteredLinkList.RemoveRange(8, _filteredLinkList.Count - 8);
                            }

                            lb.DataContext = _filteredLinkList;
                            if (lb.Items.Count > 0)
                            {
                                lb.SelectedItem = lb.Items[0];
                            }
                        }
                    }
                    else
                    {
                        _isFileExplorerMode = false;
                        lb.DataContext = null;
                    }
                }
            }
        }

        private string GetFormattedTitleFromBestResult(IList<MatchResult> matchResults)
        {
            var bestResult = matchResults.FirstOrDefault(r => r.Matched) ?? matchResults.First();
            return new XamlHighlighter().Highlight(bestResult.StringParts);
        }

        private void OnEnterPressed(object sender, ExecutedRoutedEventArgs e)
        {
            if (_isLinkQuiryMode)
            {
                Execute();
            }
            else
            {
                Switch();
            }
            e.Handled = true;
        }

        private void ListBoxItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (_isLinkQuiryMode)
            {
                Execute();
            }
            else
            {
                Switch();
            }
            e.Handled = true;
        }

        private async void CloseWindow(object sender, ExecutedRoutedEventArgs e)
        {
            if (!_isLinkQuiryMode)
            {
                var windows = lb.SelectedItems.Cast<AppWindowViewModel>().ToList();
                foreach (var win in windows)
                {
                    bool isClosed = await _windowCloser.TryCloseAsync(win);
                    if (isClosed)
                        RemoveWindow(win);
                }

                if (lb.Items.Count == 0)
                    HideWindow();

                e.Handled = true;
            }
        }

        private void RemoveWindow(AppWindowViewModel window)
        {
            int index = _filteredWindowList.IndexOf(window);
            if (index < 0)
                return;

            if (lb.SelectedIndex == index)
            {
                if (_filteredWindowList.Count > index + 1)
                    lb.SelectedIndex++;
                else
                {
                    if (index > 0)
                        lb.SelectedIndex--;
                }
            }

            _filteredWindowList.Remove(window);
            _unfilteredWindowList.Remove(window);
        }

        private void ScrollListUp(object sender, ExecutedRoutedEventArgs e)
        {
            PreviousItem();
            e.Handled = true;
        }

        private void PreviousItem()
        {
            if (lb.Items.Count > 0)
            {
                if (lb.SelectedIndex != 0)
                {
                    lb.SelectedIndex--;
                }
                else
                {
                    lb.SelectedIndex = lb.Items.Count - 1;
                }

                ScrollSelectedItemIntoView();
                UpdateTextBoxFromSelection();
            }
        }

        private void ScrollListDown(object sender, ExecutedRoutedEventArgs e)
        {
            NextItem();
            e.Handled = true;
        }

        private void NextItem()
        {
            if (lb.Items.Count > 0)
            {
                if (lb.SelectedIndex != lb.Items.Count - 1)
                {
                    lb.SelectedIndex++;
                }
                else
                {
                    lb.SelectedIndex = 0;
                }

                ScrollSelectedItemIntoView();
                UpdateTextBoxFromSelection();
            }
        }

        private void ScrollSelectedItemIntoView()
        {
            var selectedItem = lb.SelectedItem;
            if (selectedItem != null)
            {
                lb.ScrollIntoView(selectedItem);
            }
        }

        /// <summary>
        /// Update text box with selected item's path in file explorer mode
        /// </summary>
        private void UpdateTextBoxFromSelection()
        {
            if (!_isFileExplorerMode || lb.SelectedItem == null)
                return;

            var item = lb.SelectedItem as ListItemInfo;
            if (item == null)
                return;

            // Don't update for ".." (parent directory)
            if (item.FormattedTitle == "..")
                return;

            string path = item.TagData;
            if (!string.IsNullOrEmpty(path))
            {
                _suppressTextChanged = true;
                tb.Text = path;
                tb.CaretIndex = tb.Text.Length;
                _suppressTextChanged = false;
            }
        }

        private void MainWindow_OnLostFocus(object sender, EventArgs e)
        {
            HideWindow();
        }

        private void MainWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            DisableSystemMenu();
        }

        private void DisableSystemMenu()
        {
            var windowHandle = new WindowInteropHelper(this).Handle;
            var window = new SystemWindow(windowHandle);
            window.Style &= ~WindowStyleFlags.SYSMENU;
        }

        private void ShowHelpTextBlock_OnPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var duration = new Duration(TimeSpan.FromSeconds(0.150));
            var newHeight = HelpPanel.Height > 0 ? 0 : +17;
            HelpPanel.BeginAnimation(HeightProperty, new DoubleAnimation(HelpPanel.Height, newHeight, duration));
        }

        #endregion

        private enum InitialFocus {
            NextItem,
            PreviousItem
        }
    }
}