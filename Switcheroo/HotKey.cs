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

namespace Switcheroo {
    public class HotKey : ManagedWinapi.Hotkey {
        public void LoadSettings()
        {
            KeyCode = (System.Windows.Forms.Keys)Properties.Settings.Default.HotKey;
            WindowsKey = Properties.Settings.Default.WindowsKey;
            Alt = Properties.Settings.Default.Alt;
            Ctrl = Properties.Settings.Default.Ctrl;
            Shift = Properties.Settings.Default.Shift;
        }

        /// <summary>
        /// Save settings to Properties.Settings.Default without calling Save()
        /// Call Settings.Default.Save() separately after all settings are updated
        /// </summary>
        public void SaveSettings(bool saveImmediately = false)
        {
            Properties.Settings.Default.HotKey = (int)KeyCode;
            Properties.Settings.Default.WindowsKey = WindowsKey;
            Properties.Settings.Default.Alt = Alt;
            Properties.Settings.Default.Ctrl = Ctrl;
            Properties.Settings.Default.Shift = Shift;
            
            if (saveImmediately)
            {
                Properties.Settings.Default.Save();
            }
        }
    }

    public class HotKeyForExecuter : ManagedWinapi.Hotkey {
        public void LoadSettings()
        {
            KeyCode = (System.Windows.Forms.Keys)Properties.Settings.Default.LnkHotKey;
            WindowsKey = Properties.Settings.Default.LnkWindowsKey;
            Alt = Properties.Settings.Default.LnkAlt;
            Ctrl = Properties.Settings.Default.LnkCtrl;
            Shift = Properties.Settings.Default.LnkShift;
        }

        /// <summary>
        /// Save settings to Properties.Settings.Default without calling Save()
        /// Call Settings.Default.Save() separately after all settings are updated
        /// </summary>
        public void SaveSettings(bool saveImmediately = false)
        {
            Properties.Settings.Default.LnkHotKey = (int)KeyCode;
            Properties.Settings.Default.LnkWindowsKey = WindowsKey;
            Properties.Settings.Default.LnkAlt = Alt;
            Properties.Settings.Default.LnkCtrl = Ctrl;
            Properties.Settings.Default.LnkShift = Shift;
            
            if (saveImmediately)
            {
                Properties.Settings.Default.Save();
            }
        }
    }
}