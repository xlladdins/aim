﻿using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace aim
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private string alert_level = @"SOFTWARE\KALX\xll";
        private enum XLL_ALERT : uint
        {
            ERROR = 1,
            WARNING = 2,
            INFO = 4,
            // LOG = 8,
        };

        public Ribbon1()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("aim.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public Bitmap GetInfoIcon(Office.IRibbonControl control)
        {
            return SystemIcons.Information.ToBitmap();
        }
        public Bitmap GetWarningIcon(Office.IRibbonControl control)
        {
            return SystemIcons.Warning.ToBitmap();
        }
        public Bitmap GetErrorIcon(Office.IRibbonControl control)
        {
            return SystemIcons.Error.ToBitmap();
        }
        private XLL_ALERT GetAlert()
        {
            XLL_ALERT level = 0;

            using (var key = Registry.CurrentUser.OpenSubKey(alert_level))
            {
                var xal = key.GetValue("xll_alert_level");
                level = (XLL_ALERT)xal;
            }
            
            return level;
        }
        private void SetAlert(XLL_ALERT level)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(alert_level))
            {
                key.SetValue("xll_alert_level", level);
            }
        }
        public void OnAlert(Office.IRibbonControl control, bool pressed)
        {
            XLL_ALERT level = GetAlert();
            Enum.TryParse(control.Tag, out XLL_ALERT value);

            if (pressed)
            {
                level |= value;
            }
            else
            {
                level &= ~value;
            }

            SetAlert(level);
        }

        #endregion

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
