using Microsoft.Win32;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Xml;
using System.Xml.Xsl;

// TODO:  Follow these steps to enable the Ribbon (XML) item:
namespace aim
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private string alert_level = @"SOFTWARE\KALX\xll";
        private enum XLL_ALERT
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
            string ribbon_xml;
            // read from https://xlladdins.com/...
            string addin_xml = @"
<addins>
<addin>
<name>xll_math</name>
<description>Functions from the &lt;cmath&gt; library</description>
</addin>
<addin>
<name>xll_registry</name>
<description>Store and retrieve data from the Windows registry</description>
</addin>
</addins>
";
            string addin_xsl = GetResourceText("aim.Ribbon1.xsl");
            using (XmlTextReader xmlReader = new XmlTextReader(new StringReader(addin_xml)))
            {
                using (XmlTextReader xslReader = new XmlTextReader(new StringReader(addin_xsl)))
                {
                    XslCompiledTransform xslTransform = new XslCompiledTransform();
                    xslTransform.Load(xslReader);
                    using (StringWriter xmlWriter = new StringWriter())
                    {
                        xslTransform.Transform(xmlReader, null, xmlWriter);
                        ribbon_xml = xmlWriter.ToString();
                    }
                }

            }

            return ribbon_xml;
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
                var xal = key.GetValue("xll_alert_level", 0);
                level = (XLL_ALERT)Enum.ToObject(typeof(XLL_ALERT), xal);
            }
            
            return level;
        }
        private void SetAlert(XLL_ALERT level)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(alert_level))
            {
                key.SetValue("xll_alert_level", level, RegistryValueKind.DWord);
            }
        }
        public bool GetPressedAlert(Office.IRibbonControl control)
        {
            XLL_ALERT level = GetAlert();
            Enum.TryParse(control.Tag, out XLL_ALERT value);

            return level.HasFlag(value);
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
        public void OnAddIn(Office.IRibbonControl control, bool pressed)
        {

        }
        public bool GetPressedAddIn(Office.IRibbonControl control)
        {
            return true;
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
