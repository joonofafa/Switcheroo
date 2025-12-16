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
using Switcheroo.Core;
using Switcheroo.Properties;
using System.Text;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;
using Application = System.Windows.Application;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MessageBox = System.Windows.MessageBox;

namespace Switcheroo {
    public partial class OptionsWindow : Window {
        private readonly HotKey _hotkey;
        private readonly HotKeyForExecuter _lnkHotkey;
        private HotkeyViewModel _hotkeyViewModel;
        private HotkeyViewModel _lnkHotkeyViewModel;

        public OptionsWindow()
        {
            InitializeComponent();

            // Show what's already selected     
            _hotkey = (HotKey)Application.Current.Properties["hotkey"];
            _lnkHotkey = (HotKeyForExecuter)Application.Current.Properties["lnkHotkey"];

            try
            {
                _hotkey.LoadSettings();
                _lnkHotkey.LoadSettings();
            }
            catch (HotkeyAlreadyInUseException)
            {
            }

            _hotkeyViewModel = new HotkeyViewModel
            {
                KeyCode = KeyInterop.KeyFromVirtualKey((int)_hotkey.KeyCode),
                Alt = _hotkey.Alt,
                Ctrl = _hotkey.Ctrl,
                Windows = _hotkey.WindowsKey,
                Shift = _hotkey.Shift
            };
            _lnkHotkeyViewModel = new HotkeyViewModel
            {
                KeyCode = KeyInterop.KeyFromVirtualKey((int)_lnkHotkey.KeyCode),
                Alt = _lnkHotkey.Alt,
                Ctrl = _lnkHotkey.Ctrl,
                Windows = _lnkHotkey.WindowsKey,
                Shift = _lnkHotkey.Shift
            };

            HotKeyCheckBox.IsChecked = Settings.Default.EnableHotKey;
            LnkHotKeyCheckBok.IsChecked = Settings.Default.EnableLnkHotkey;
            HotkeyPreview.Text = _hotkeyViewModel.ToString();
            LnkHotkeyPreview.Text = _lnkHotkeyViewModel.ToString();
            HotkeyPreview.IsEnabled = Settings.Default.EnableHotKey;
            LnkHotkeyPreview.IsEnabled = Settings.Default.EnableLnkHotkey;
            AltTabCheckBox.IsChecked = Settings.Default.AltTabHook;
            AutoSwitch.IsChecked = Settings.Default.AutoSwitch;
            AutoSwitch.IsEnabled = Settings.Default.AltTabHook;
            RunAsAdministrator.IsChecked = Settings.Default.RunAsAdmin;
            
            // Initialize File Explorer Max Items ComboBox (10 to 100, step 5)
            for (int i = 10; i <= 100; i += 5)
            {
                FileExplorerMaxItemsComboBox.Items.Add(i);
            }
            
            // Set current value
            int currentMaxItems = Settings.Default.FileExplorerMaxItems;
            if (currentMaxItems < 10) currentMaxItems = 10;
            if (currentMaxItems > 100) currentMaxItems = 100;
            // Round to nearest 5
            currentMaxItems = ((currentMaxItems + 2) / 5) * 5;
            FileExplorerMaxItemsComboBox.SelectedItem = currentMaxItems;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var closeOptionsWindow = true;

            try
            {
                // Step 1: Disable both hotkeys first to prevent conflicts
                _hotkey.Enabled = false;
                _lnkHotkey.Enabled = false;

                // Step 2: Get new enable states from UI (use these for logic below)
                bool enableHotKey = HotKeyCheckBox.IsChecked.GetValueOrDefault();
                bool enableLnkHotkey = LnkHotKeyCheckBok.IsChecked.GetValueOrDefault();

                // Step 3: Update hotkey properties (but don't enable yet)
                _hotkey.Alt = _hotkeyViewModel.Alt;
                _hotkey.Shift = _hotkeyViewModel.Shift;
                _hotkey.Ctrl = _hotkeyViewModel.Ctrl;
                _hotkey.WindowsKey = _hotkeyViewModel.Windows;
                _hotkey.KeyCode = (Keys)KeyInterop.VirtualKeyFromKey(_hotkeyViewModel.KeyCode);

                _lnkHotkey.Alt = _lnkHotkeyViewModel.Alt;
                _lnkHotkey.Shift = _lnkHotkeyViewModel.Shift;
                _lnkHotkey.Ctrl = _lnkHotkeyViewModel.Ctrl;
                _lnkHotkey.WindowsKey = _lnkHotkeyViewModel.Windows;
                _lnkHotkey.KeyCode = (Keys)KeyInterop.VirtualKeyFromKey(_lnkHotkeyViewModel.KeyCode);

                // Step 4: Save all settings to Settings.Default (without calling Save() yet)
                _hotkey.SaveSettings(saveImmediately: false);
                _lnkHotkey.SaveSettings(saveImmediately: false);
                
                Settings.Default.EnableHotKey = enableHotKey;
                Settings.Default.EnableLnkHotkey = enableLnkHotkey;
                Settings.Default.AltTabHook = AltTabCheckBox.IsChecked.GetValueOrDefault();
                Settings.Default.AutoSwitch = AutoSwitch.IsChecked.GetValueOrDefault();
                Settings.Default.RunAsAdmin = RunAsAdministrator.IsChecked.GetValueOrDefault();
                
                // Save File Explorer Max Items
                if (FileExplorerMaxItemsComboBox.SelectedItem != null)
                {
                    Settings.Default.FileExplorerMaxItems = (int)FileExplorerMaxItemsComboBox.SelectedItem;
                }

                // Step 5: Save all settings at once
                Settings.Default.Save();

                // Step 6: Now enable hotkeys based on new settings
                if (enableHotKey)
                {
                    _hotkey.Enabled = true;
                }

                if (enableLnkHotkey)
                {
                    _lnkHotkey.Enabled = true;
                }
            }
            catch (HotkeyAlreadyInUseException)
            {
                var boxText = "Sorry! The selected shortcut for activating Switcheroo is in use by another program. " +
                              "Please choose another.";
                MessageBox.Show(boxText, "Shortcut already in use", MessageBoxButton.OK, MessageBoxImage.Warning);
                closeOptionsWindow = false;
            }

            if (closeOptionsWindow)
            {
                Close();
            }
        }

        private class HotkeyViewModel {
            public Key KeyCode { get; set; }
            public bool Shift { get; set; }
            public bool Alt { get; set; }
            public bool Ctrl { get; set; }
            public bool Windows { get; set; }

            public override string ToString()
            {
                var shortcutText = new StringBuilder();

                if (Ctrl)
                {
                    shortcutText.Append("Ctrl + ");
                }

                if (Shift)
                {
                    shortcutText.Append("Shift + ");
                }

                if (Alt)
                {
                    shortcutText.Append("Alt + ");
                }

                if (Windows)
                {
                    shortcutText.Append("Win + ");
                }

                var keyString =
                    KeyboardHelper.CodeToString((uint)KeyInterop.VirtualKeyFromKey(KeyCode)).ToUpper().Trim();
                if (keyString.Length == 0)
                {
                    keyString = new KeysConverter().ConvertToString(KeyCode);
                }

                // If the user presses "Escape" then show "Escape" :)
                if (keyString == "\u001B")
                {
                    keyString = "Escape";
                }

                shortcutText.Append(keyString);
                return shortcutText.ToString();
            }
        }

        private void HotkeyPreview_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            // The text box grabs all input
            e.Handled = true;

            // Fetch the actual shortcut key
            var key = e.Key == Key.System ? e.SystemKey : e.Key;

            // Ignore modifier keys
            if (key == Key.LeftShift || key == Key.RightShift
                || key == Key.LeftCtrl || key == Key.RightCtrl
                || key == Key.LeftAlt || key == Key.RightAlt
                || key == Key.LWin || key == Key.RWin)
            {
                return;
            }

            var previewHotkeyModel = new HotkeyViewModel();
            previewHotkeyModel.Ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
            previewHotkeyModel.Shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
            previewHotkeyModel.Alt = (Keyboard.Modifiers & ModifierKeys.Alt) != 0;

            var winLKey = new KeyboardKey(Keys.LWin);
            var winRKey = new KeyboardKey(Keys.RWin);
            previewHotkeyModel.Windows = (winLKey.State & 0x8000) == 0x8000 || (winRKey.State & 0x8000) == 0x8000;
            previewHotkeyModel.KeyCode = key;

            var previewText = previewHotkeyModel.ToString();

            // Jump to the next element if the user presses only the Tab key
            if (previewText == "Tab")
            {
                ((UIElement)sender).MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                return;
            }

            HotkeyPreview.Text = previewText;
            _hotkeyViewModel = previewHotkeyModel;
        }
        private void HotkeyPreview_OnGotFocus(object sender, RoutedEventArgs e)
        {
            // Disable the current hotkey while the hotkey field is active
            _hotkey.Enabled = false;
        }

        private void HotkeyPreview_OnLostFocus(object sender, RoutedEventArgs e)
        {
            // Don't re-enable here - let Ok_Click handle enabling with new settings
            // Re-enabling with old settings can cause conflicts
            // Only re-enable if the checkbox is still checked (for Cancel case)
            if (!HotKeyCheckBox.IsChecked.GetValueOrDefault())
            {
                return;
            }

            try
            {
                // Temporarily apply the preview settings before re-enabling
                _hotkey.Alt = _hotkeyViewModel.Alt;
                _hotkey.Shift = _hotkeyViewModel.Shift;
                _hotkey.Ctrl = _hotkeyViewModel.Ctrl;
                _hotkey.WindowsKey = _hotkeyViewModel.Windows;
                _hotkey.KeyCode = (Keys)KeyInterop.VirtualKeyFromKey(_hotkeyViewModel.KeyCode);
                _hotkey.Enabled = true;
            }
            catch (HotkeyAlreadyInUseException)
            {
                // It is alright if the hotkey can't be reactivated.
            }
        }

        private void LnkHotkeyPreview_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            // The text box grabs all input
            e.Handled = true;

            // Fetch the actual shortcut key
            var key = e.Key == Key.System ? e.SystemKey : e.Key;

            // Ignore modifier keys
            if (key == Key.LeftShift || key == Key.RightShift
                || key == Key.LeftCtrl || key == Key.RightCtrl
                || key == Key.LeftAlt || key == Key.RightAlt
                || key == Key.LWin || key == Key.RWin)
            {
                return;
            }

            var previewHotkeyModel = new HotkeyViewModel();
            previewHotkeyModel.Ctrl = (Keyboard.Modifiers & ModifierKeys.Control) != 0;
            previewHotkeyModel.Shift = (Keyboard.Modifiers & ModifierKeys.Shift) != 0;
            previewHotkeyModel.Alt = (Keyboard.Modifiers & ModifierKeys.Alt) != 0;

            var winLKey = new KeyboardKey(Keys.LWin);
            var winRKey = new KeyboardKey(Keys.RWin);
            previewHotkeyModel.Windows = (winLKey.State & 0x8000) == 0x8000 || (winRKey.State & 0x8000) == 0x8000;
            previewHotkeyModel.KeyCode = key;

            var previewText = previewHotkeyModel.ToString();

            // Jump to the next element if the user presses only the Tab key
            if (previewText == "Tab")
            {
                ((UIElement)sender).MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                return;
            }

            LnkHotkeyPreview.Text = previewText;
            _lnkHotkeyViewModel = previewHotkeyModel;
        }

        private void LnkHotkeyPreview_OnGotFocus(object sender, RoutedEventArgs e)
        {
            // Disable the current hotkey while the hotkey field is active
            _lnkHotkey.Enabled = false;
        }

        private void LnkHotkeyPreview_OnLostFocus(object sender, RoutedEventArgs e)
        {
            // Don't re-enable here - let Ok_Click handle enabling with new settings
            // Re-enabling with old settings can cause conflicts
            // Only re-enable if the checkbox is still checked (for Cancel case)
            if (!LnkHotKeyCheckBok.IsChecked.GetValueOrDefault())
            {
                return;
            }

            try
            {
                // Temporarily apply the preview settings before re-enabling
                _lnkHotkey.Alt = _lnkHotkeyViewModel.Alt;
                _lnkHotkey.Shift = _lnkHotkeyViewModel.Shift;
                _lnkHotkey.Ctrl = _lnkHotkeyViewModel.Ctrl;
                _lnkHotkey.WindowsKey = _lnkHotkeyViewModel.Windows;
                _lnkHotkey.KeyCode = (Keys)KeyInterop.VirtualKeyFromKey(_lnkHotkeyViewModel.KeyCode);
                _lnkHotkey.Enabled = true;
            }
            catch (HotkeyAlreadyInUseException)
            {
                // It is alright if the hotkey can't be reactivated.
            }
        }

        private void AltTabCheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            AutoSwitch.IsEnabled = true;
        }

        private void AltTabCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
        {
            AutoSwitch.IsEnabled = false;
            AutoSwitch.IsChecked = false;
        }

        private void HotKeyCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            HotkeyPreview.IsEnabled = true;
        }

        private void HotKeyCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
        {
            HotkeyPreview.IsEnabled = false;
        }

        private void LnkHotKeyCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            LnkHotkeyPreview.IsEnabled = true;
        }

        private void LnkHotKeyCheckBox_OnUnchecked(object sender, RoutedEventArgs e)
        {
            LnkHotkeyPreview.IsEnabled = false;
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            // Reload settings from file
            Settings.Default.Reload();
            
            // Get the main window and refresh its settings
            var mainWindow = Application.Current.MainWindow as MainWindow;
            mainWindow?.RefreshSettings();
            
            MessageBox.Show("Settings and configuration files have been refreshed.", 
                "Refresh Complete", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}