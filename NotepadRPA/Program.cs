using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace NotepadRPA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            bool isNotepadRunning()
            {
                Process[] allProcesses = Process.GetProcesses();

                for (int i = 0; i < allProcesses.Length; i++)
                {
                    if (allProcesses[i].ProcessName == "notepad")
                    {
                        return true;
                    }
                }
                return false;
            }

            Console.WriteLine("Hello World!");
            if (isNotepadRunning())
            {
                Debug.WriteLine("Notepad is running");
            }
            else
            {
                Debug.WriteLine("Notepad is not running");
                Process.Start("notepad.exe");
            }


            //[DllImport("User32")]
        }
    }
}
