using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Office = NetOffice.OfficeApi;

namespace InnerAddin
{
    internal class MyRibbonExtensibility : NetOffice.OfficeApi.Native.IRibbonExtensibility
    {
        public string GetCustomUI(string RibbonID)
        {
            MessageBox.Show("getCustomUI");
            var result = ReadString("RibbonUI.xml", typeof(MyRibbonExtensibility).Assembly);
            return result;
        }

        public void SampleButton_Click(Office.IRibbonControl control)
        {
            MessageBox.Show("Thanks!", "InnerAddin.MyRibbonExtensibility");
        }

        public string ReadString(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = ReadStream(resourceAddress, assembly);
            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(resourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource string."));

            string text = textStreamReader.ReadToEnd();
            textStreamReader.Close();
            resourceStream.Close();
            return text;
        }

        public Stream ReadStream(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = assembly.GetManifestResourceStream(resourceAddress);
            if (resourceStream == null)
            {
                string target = typeof(MyRibbonExtensibility).Namespace + "." + resourceAddress;
                MessageBox.Show(target);
                resourceStream = assembly.GetManifestResourceStream(target);
            }

            if (resourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            return resourceStream;
        }
    }
}
