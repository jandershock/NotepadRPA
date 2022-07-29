using System.Threading;
using System.Windows.Automation;

namespace NotepadRPA.Helpers
{
    internal class ElementFunctions
    {
        internal static void ExpandElement(AutomationElement element)
        {
            ExpandCollapsePattern expandCollapsePattern = (ExpandCollapsePattern)element.GetCurrentPattern(ExpandCollapsePattern.Pattern);
            expandCollapsePattern.Expand();
        }

        internal static void InvokeElement(AutomationElement element)
        {
            InvokePattern invokePattern = (InvokePattern)element.GetCurrentPattern(InvokePattern.Pattern);
            invokePattern.Invoke();
        }

        internal static void InvokeElementBackgroundProcess(AutomationElement element)
        {
            InvokePattern invokePattern = (InvokePattern)element.GetCurrentPattern(InvokePattern.Pattern);
            Thread background = new Thread(invokePattern.Invoke);
            background.IsBackground = true;
            background.Start();
        }

        internal static void SetValue(AutomationElement element, string valueString)
        {
            ValuePattern valuePattern = (ValuePattern)element.GetCurrentPattern(ValuePattern.Pattern);
            valuePattern.SetValue(valueString);
        }
    }
}
