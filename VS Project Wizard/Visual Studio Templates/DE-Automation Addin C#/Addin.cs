using System;
using System.Reflection;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;
$usingItems$
namespace $safeprojectname$
{
    [ComVisible(true)]
    [GuidAttribute("$randomGuid$"), ProgId("$safeprojectname$.$safeitemname$")]
    public class Addin :IDTExtensibility2$ribbonImplement$
    {
        public Addin()
        {
            /* Initialize NetOffice */
            LateBindingApi.Core.Factory.Initialize();
        }
        
        #region IDTExtensibility2 Members

 	void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
$classicUICreateCall$            
        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
           
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
$classicUIRemoveCall$
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
           
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
             
        }

        #endregion
$ribbonUIImplementMethod$$classicUICreateRemoveMethod$
        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Addin));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();
                
$registerCode$
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register $safeitemname$", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            { 
                // unregister addin
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                
                // unregister addin in office
$unregisterCode$
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister $safeitemname$", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
$helperCode$
    }
}
