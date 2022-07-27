using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;

namespace NotepadRPA
{
    internal class Program
    {
        internal struct Windowplacement
        {
            internal int length;
            internal int flags;
            internal int showCmd;
            internal System.Drawing.Point ptMinPosition;
            internal System.Drawing.Point ptMaxPosition;
            internal System.Drawing.Rectangle rcNormalPosition;
        }

        static void Main(string[] args)
        {
            [DllImport("user32")]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);

            [DllImport("user32")]
            static extern bool SetForegroundWindow(IntPtr hwnd);

            [DllImport("user32")]
            [return: MarshalAs(UnmanagedType.Bool)]
            static extern bool GetWindowPlacement(IntPtr hwnd, ref Windowplacement lpwndpl);

            bool isNotepadRunning()
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

            AutomationElement FindTreeViewDescendants(AutomationElement targetTreeViewElement, string searchParam)
            {

                if (targetTreeViewElement == null)
                    return null;

                AutomationElement foundNode = null;
                AutomationElement elementNode =
                    TreeWalker.ControlViewWalker.GetFirstChild(targetTreeViewElement);

                while (elementNode != null)
                {
                    string controlName =
                        (elementNode.Current.Name == "") ?
                        "Unnamed control" : elementNode.Current.Name;
                    string autoIdName =
                        (elementNode.Current.AutomationId == "") ?
                        "No AutomationID" : elementNode.Current.AutomationId;
                    Debug.WriteLine($"Name: {controlName} AutomationID: {autoIdName}");

                    if (controlName == searchParam)
                    {
                        foundNode = elementNode;
                        return foundNode;
                    }

                    foundNode = FindTreeViewDescendants(elementNode, searchParam);
                    if (foundNode == null)
                    {
                        elementNode = TreeWalker.ControlViewWalker.GetNextSibling(elementNode);
                    }
                    else
                    {
                        return foundNode;
                    }

                }
                return foundNode;
            }

            void ExpandElement(AutomationElement element)
            {
                ExpandCollapsePattern expandCollapsePattern = (ExpandCollapsePattern)element.GetCurrentPattern(ExpandCollapsePattern.Pattern);
                expandCollapsePattern.Expand();
            }

            void InvokeElement(AutomationElement element)
            {
                InvokePattern invokePattern = (InvokePattern)element.GetCurrentPattern(InvokePattern.Pattern);
                invokePattern.Invoke();
            }

            if (isNotepadRunning())
            {
                Debug.WriteLine("Please close Notepad before running this program.");

                Environment.Exit(0);
            }
            Debug.WriteLine("Notepad is not running");
            // Start notepad and wait for idle state
            Process np = Process.Start("notepad.exe");
            np.WaitForInputIdle();

            //Get notepad window handle and recursively examine all elements in its associated tree
            IntPtr windowHandle = np.MainWindowHandle;
            AutomationElement notepadElement = AutomationElement.FromHandle(windowHandle);
            AutomationElement fileElement = FindTreeViewDescendants(notepadElement, "File");

            //Click on "File" dropdown
            ExpandElement(fileElement);

            //Find the "Save As..." element and click it
            AutomationElement saveAsElement = FindTreeViewDescendants(fileElement, "Save As...");
            InvokeElement(saveAsElement);

            Debug.WriteLine("");
            //Get element tree for save as dialog menu
            AutomationElement saveAsDialogElement = FindTreeViewDescendants(AutomationElement.RootElement, "Save As");

            Debug.WriteLine("");
            AutomationElement search = FindTreeViewDescendants(saveAsDialogElement, "");

            //Close notepad
            np.CloseMainWindow();
        }
    }
}
