using System;
using System.Reflection;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;
$usingItems$
namespace $safeprojectname$
{
    [GuidAttribute("$randomGuid$"), ProgId("$safeprojectname$.$safeitemname$"), ComVisible(true)]
    public class Addin :IDTExtensibility2$ribbonImplement$
    {
$ApplicationField$
        public Addin()
        {
        }
        
        #region IDTExtensibility2 Members

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
$ApplicationConstruction$

 		// If the addin not connected during startup, we call OnStartupComplete at hand
    		if (ConnectMode != ext_ConnectMode.ext_cm_Startup)
        		OnStartupComplete(ref custom);
        }

 	public void OnStartupComplete(ref Array custom)
        {
$classicUICreateCall$            
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
 		// If this is not because of host shutdown(removed by user for example) we call OnBeginShutdown at hand
    		if (RemoveMode != ext_DisconnectMode.ext_dm_HostShutdown)
        		OnBeginShutdown(ref custom);

$ApplicationDestroy$
        }

        public void OnBeginShutdown(ref Array custom)
        {
$classicUIRemoveCall$             
        }

        public void OnAddInsUpdate(ref Array custom)
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
