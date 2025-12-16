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

using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace Switcheroo {
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            // Handle exceptions thrown on the UI thread
            DispatcherUnhandledException += App_DispatcherUnhandledException;
            
            // Handle exceptions thrown on background threads
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            
            // Handle exceptions thrown in async methods
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;
        }

        private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[Switcheroo] UI Thread Exception: {e.Exception.Message}");
            System.Diagnostics.Debug.WriteLine($"[Switcheroo] Stack Trace: {e.Exception.StackTrace}");
            
            // Mark as handled to prevent app crash
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Switcheroo] AppDomain Exception: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[Switcheroo] Stack Trace: {ex.StackTrace}");
            }
        }

        private void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[Switcheroo] Task Exception: {e.Exception.Message}");
            System.Diagnostics.Debug.WriteLine($"[Switcheroo] Stack Trace: {e.Exception.StackTrace}");
            
            // Mark as observed to prevent app crash
            e.SetObserved();
        }
    }
}