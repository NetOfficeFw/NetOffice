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
                        Utils.Dialog.ShowMessageBox("This is the first sample button. " + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinCS4", DialogResult.None);
                        break;
                    case "customButton2":
                        Utils.Dialog.ShowMessageBox("This is the second sample button. " + Application.FriendlyTypeName, "NetOfficeTools.SuperAddinCS4", DialogResult.None);
                        break;
                    default:
                        Utils.Dialog.ShowMessageBox("Unkown Control Id: " + control.Id, "NetOfficeTools.SuperAddinCS4", DialogResult.None);
                        break;
                }
            }
            catch (Exception throwedException)
            {
                Utils.Dialog.ShowError(throwedException, "Unexpected state in SuperAddinCS4 OnAction");
            }
        }

        #endregion

        #region Error Handler

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "Unexpected state in SuperAddinCS4 " + methodKind.ToString());
        }

        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An register error occurend in " + methodKind.ToString(), "SuperAddinCS4");
        }

        #endregion
    }
}
