using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;

namespace NotepadRPA.Helpers
{
    internal class Utils
    {
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

        [DllImport("user32")]
        static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowPlacement(IntPtr hwnd, ref Windowplacement lpwndpl);

        internal static bool isNotepadRunning()
        {
            Process[] allProcesses = Process.GetProcesses();

            for (int i = 0; i < allProcesses.Length; i++)
            {
                if (allProcesses[i].ProcessName == "notepad")
                {
                    MessageBox.Show("Please close Notepad before running this program.", "Notepad Process Detected", MessageBoxButton.OK, MessageBoxImage.Error);

                    IntPtr notepad_hwnd = allProcesses[i].MainWindowHandle;
                    Windowplacement placement = new Windowplacement();

                    GetWindowPlacement(notepad_hwnd, ref placement);
                    if (placement.showCmd == 2)
                    {
                        ShowWindow(notepad_hwnd, 9); // 9 is the flag value for restore
                    }
                    SetForegroundWindow(notepad_hwnd);

                    return true;
                }
            }
            return false;
        }
    }
}