using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using MSO = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
//using Microsoft.Office.Tools.Ribbon;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ComposeRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace Ctse.Outlook.Toolbox.Ribbons
{
    [ComVisible(true)]
    public class CustomRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        //private readonly Microsoft.Office.Interop.Outlook.Application _applicationObject;

        public CustomRibbon()
        {
            //this._applicationObject = new Microsoft.Office.Interop.Outlook.Application();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            string str = string.Empty;
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    break;
                case "Microsoft.Outlook.Mail.Read":
                    str = GetResourceText("Ctse.Outlook.Toolbox.Ribbons.ReadRibbon.xml");
                    break;
                case "Microsoft.Outlook.Mail.Compose":
                    str = GetResourceText("Ctse.Outlook.Toolbox.Ribbons.ComposeRibbon.xml");
                    break;
            }
            return str;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /*
        public void OnAction(object sender, RibbonControlEventArgs e)
        {
            if (e.Control.Id == "ExportEML")
            {

            }
        }
        */

        public void OnAction(Office.IRibbonControl control)//, bool isPressed)
        {
            //Globals.ThisAddIn.Application.ActiveExplorer().Selection
            var msg = control.Context.CurrentItem as MSO.MailItem;
            if (control.Id == "ExportEMLTemplate")
            {
                Helper.SaveMessageToEml(msg, true);
            }
            else if (control.Id == "ExportEMLEmail")
            {
                Helper.SaveMessageToEml(msg, false);
            }
        }

        #endregion

        public bool GetVisible(Office.IRibbonControl control)
        {
            string id = control.Id;
            bool flag = true;
            /*
            switch (id)
            {
                case "ExportEML":
                    flag = LocalPreferences.Instance.ExportEnabled;
                    break;
            }
            */
            return flag;
        }

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
