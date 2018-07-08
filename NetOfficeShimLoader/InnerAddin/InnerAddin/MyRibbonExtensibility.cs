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
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public class MyRibbonExtensibility : NetOffice.OfficeApi.Native.IRibbonExtensibility
    {
        private Guid _id = Guid.NewGuid();

        public MyRibbonExtensibility(Connect parent)
        {
            Parent = parent;
        }

        private Connect Parent { get; set; }

        public void SampleButton1_Click(Office.IRibbonControl control)
        {
            try
            {
                //MessageBox.Show("Me is " + _id.ToString());

                if (Parent.HasShimHost)
                {
                    Parent.CallShimHost();
                }
                else
                {
                    string message = "I dont have a shim host :(";
                    MessageBox.Show(message, "InnerAddin.MyRibbonExtensibility");
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString(), "SampleButton1_Click");
                throw;
            }
        }

        public void SampleButton2_Click(Office.IRibbonControl control)
        {
            string message = "I'm alive :)" + Environment.NewLine + "My Hash is: " + GetHashCode();
            MessageBox.Show(message, "InnerAddin.MyRibbonExtensibility");
        }

        public string GetCustomUI(string RibbonID)
        {
            //MessageBox.Show("getCustomUI");
            var result = ReadString("RibbonUI.xml", typeof(MyRibbonExtensibility).Assembly);
            return result;
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
                //MessageBox.Show(target);
                resourceStream = assembly.GetManifestResourceStream(target);
            }

            if (resourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            return resourceStream;
        }
    }
}
