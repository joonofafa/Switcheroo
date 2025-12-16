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
            if (!tb.IsEnabled)
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

        // Windows API for simulating key press using SendInput
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public INPUTUNION u;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)]
        private struct INPUTUNION
        {
            [System.Runtime.InteropServices.FieldOffset(0)]
            public KEYBDINPUT ki;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private const uint INPUT_KEYBOARD = 1;
        private const ushort VK_SNAPSHOT = 0x2C;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        private void TextBox_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // Handle PrintScreen key - prevent TextBox from processing it
            if (e.Key == Key.PrintScreen || e.Key == Key.Snapshot)
            {
                e.Handled = true;
                
                // Hide window and release focus before sending PrintScreen
                Opacity = 0;
                Hide();
                
                // Execute PrintScreen after window is fully hidden
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    System.Threading.Thread.Sleep(100); // Brief delay to ensure window is hidden
                    
                    // Send PrintScreen key using SendInput
                    INPUT[] inputs = new INPUT[2];
                    
                    // Key down
                    inputs[0].type = INPUT_KEYBOARD;
                    inputs[0].u.ki.wVk = VK_SNAPSHOT;
                    inputs[0].u.ki.wScan = 0;
                    inputs[0].u.ki.dwFlags = 0;
                    inputs[0].u.ki.time = 0;
                    inputs[0].u.ki.dwExtraInfo = IntPtr.Zero;
                    
                    // Key up
                    inputs[1].type = INPUT_KEYBOARD;
                    inputs[1].u.ki.wVk = VK_SNAPSHOT;
                    inputs[1].u.ki.wScan = 0;
                    inputs[1].u.ki.dwFlags = KEYEVENTF_KEYUP;
                    inputs[1].u.ki.time = 0;
                    inputs[1].u.ki.dwExtraInfo = IntPtr.Zero;
                    
                    SendInput(2, inputs, System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT)));
                    
                    // Restore window after a delay
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        Show();
                        Activate();
                        Opacity = 1;
                        tb.Focus();
                    }), DispatcherPriority.Background);
                }), DispatcherPriority.Input);
            }
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Block any text input from PrintScreen key (may appear as special character)
            if (e.Text.Length == 1)
            {
                char c = e.Text[0];
                // Block non-printable control characters (except common ones like space)
                if (char.IsControl(c) && c != ' ' && c != '\t' && c != '\r' && c != '\n')
                {
                    e.Handled = true;
                }
            }
        }

        #endregion

        private enum InitialFocus {
            NextItem,
            PreviousItem
        }
    }
}