using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

namespace NetOfficeTools.SuperAddinCS4
{
    [COMAddin("NetOfficeTools Super Addin Sample", "This Addin shows you how i can create a NO Tools based Addin and support multiple office products", 3)]
    [RegistryLocation(RegistrySaveLocation.CurrentUser), CustomUI("RibbonUI.xml", true)]
    [Guid("CF0E2618-37D5-4efb-BD25-58301228ED0E"), ProgId("NOToolsSuperAddinCS4.Addin"), Tweak(true)]
    [MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.PowerPoint, RegisterIn.Outlook, RegisterIn.Access, RegisterIn.MSProject)]  // visio is not supported because visio doesnt use the common office core
    public class Addin : COMAddin 
    {
        #region Ribbon UI Trigger

        public void OnAction(NetOffice.OfficeApi.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "customButton1":

                        MessageBox.Show("This is the first sample button. " + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinCS4");
                        break;
                    case "customButton2":
                        MessageBox.Show("This is the second sample button." + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinCS4");
                        break;
                    default:
                        MessageBox.Show("Unkown Control Id: " + control.Id, "NetOfficeTools.SuperAddinCS4");
                        break;
                }
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured in OnAction." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
