using NotepadRPA.Helpers;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace NotepadRPA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fileNameAndPath = Path.GetTempPath() + "hello.txt";

            //Check for running Notepad processes
            Console.WriteLine("Checking for running Notepad processes . . .");
            if (Utils.isNotepadRunning())
            {
                Console.WriteLine("Please close Notepad before running this program.");

                Environment.Exit(0);
            }

            // Start notepad and wait for idle state
            Console.WriteLine("Starting Notepad . . .");
            Process np = Process.Start("notepad.exe");
            np.WaitForInputIdle();
            Console.WriteLine("Writing to file . . .");
            SendKeys.SendWait("Hello World");

            //Get notepad window handle and recursively examine all elements in its associated tree
            Console.WriteLine("Searching for 'File' element . . .");
            IntPtr windowHandle = np.MainWindowHandle;
            AutomationElement notepadElement = AutomationElement.FromHandle(windowHandle);
            AutomationElement fileElement = notepadElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, "File"));

            //Click on "File" dropdown
            Console.WriteLine("Expanding 'File' element . . .");
            ElementFunctions.ExpandElement(fileElement);

            //Find the "Save As..." element and click it
            Console.WriteLine("Searching for 'Save As...' element . . .");
            AutomationElement saveAsElement = fileElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, "Save As..."));
            Console.WriteLine("Invoking 'Save As...' element . . .");
            ElementFunctions.InvokeElementBackgroundProcess(saveAsElement);

            //Find the "Save As" dialog window
            Console.WriteLine("Searching for 'Save As' dialog window . . .");
            //Get element tree for save as dialog menu
            Thread.Sleep(1000);
            AutomationElement saveAsDialogElement = AutomationElement.RootElement.FindFirst(TreeScope.Subtree, new AndCondition(new System.Windows.Automation.Condition[]
            {
                        new PropertyCondition(AutomationElement.NameProperty, "Save As"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
            }));

            //Find the "File name:" edit element
            Console.WriteLine("Searching for 'File name:' edit element . . .");
            PropertyCondition p1 = new PropertyCondition(AutomationElement.NameProperty, "File name:");
            PropertyCondition p2 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
            Condition[] conditionArray = new Condition[] { p1, p2 };

            AutomationElement saveLabelElement = null;
            try
            {
                saveLabelElement = saveAsDialogElement.FindFirst(TreeScope.Subtree, new AndCondition(conditionArray));
            }
            catch (NullReferenceException ex)
            {
                Console.Write("Error: Accessed element is null.");
                return;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }

            //Modify value of the file name
            Console.WriteLine("Modifying file name . . .");
            ElementFunctions.SetValue(saveLabelElement, fileNameAndPath);

            //Get Save button and invoke element
            Console.WriteLine("Searching for 'Save' button element . . .");
            AutomationElement saveButtonElement = saveAsDialogElement.FindFirst(TreeScope.Subtree, new AndCondition(new System.Windows.Automation.Condition[]
            {
                        new PropertyCondition(AutomationElement.NameProperty, "Save"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
            }));
            Console.WriteLine("Invoking 'Save' button . . .");
            ElementFunctions.InvokeElementBackgroundProcess(saveButtonElement);

            //Check if "Confirm Save As" dialog window exists
            Console.WriteLine(@"Looking for 'Confirm Save As' dialog window");
            AutomationElement confirmSaveAsDialogElement = AutomationElement.RootElement.FindFirst(TreeScope.Subtree, new AndCondition(new System.Windows.Automation.Condition[]
            {
                        new PropertyCondition(AutomationElement.NameProperty, "Confirm Save As"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
            }));

            //File already exists, confirm overwrite
            if (confirmSaveAsDialogElement != null)
            {
                //Get "Yes" button from "Confirm Save As
                Console.WriteLine("File already exists, searching for 'Yes' button . . .");
                AutomationElement yesButtonElement = confirmSaveAsDialogElement.FindFirst(TreeScope.Descendants, new AndCondition(new System.Windows.Automation.Condition[]
                {
                        new PropertyCondition(AutomationElement.NameProperty, "Yes"),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button)
                }));
                Console.WriteLine("Invoking 'Yes' button . . .");
                ElementFunctions.InvokeElement(yesButtonElement);
            }

            //Verify file was saved to location
            Console.WriteLine(File.Exists(fileNameAndPath) ? "File saved, task completed successfully" : "Warning: File was not saved");

            //Close notepad
            Console.WriteLine("Closing Notepad . . .");
            Thread.Sleep(1000);
    
            np.CloseMainWindow();

            //Type any button to exit program
            Console.WriteLine("\nPress any button to exit the program");
            Console.ReadKey();
            return;
        }
    }
}
