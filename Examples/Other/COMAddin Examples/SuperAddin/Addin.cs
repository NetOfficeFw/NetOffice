using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;

/*
    Multi-host addin example
*/
namespace SuperAddinCS4
{
    [COMAddin("Super Addin Sample CS4", "Multi-host addin example", LoadBehavior.LoadAtStartup), Codebase]
    [CustomUI("RibbonUI.xml", true), RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [ProgId("SuperAddinCS4.Connect"), Guid("CF0E2618-37D5-4efb-BD25-58301228ED0E")]
    [MultiRegister(RegisterIn.Excel, RegisterIn.Word, RegisterIn.Outlook, RegisterIn.PowerPoint, RegisterIn.Access)]
    public class Addin : COMAddin 
    {
        public void OnRibbonButtonClick(NetOffice.OfficeApi.IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "customButton1":
                        MessageBox.Show(String.Format("Hosted in {0}", Application.InstanceFriendlyName));
                        break;
                    case "customButton2":
                        MessageBox.Show(String.Format("Loading Time {0}", LoadingTimeElapsed));
                        break;
                }
            }
            catch (Exception throwedException)
            {
                Utils.Dialog.ShowError(throwedException, "Unexpected state in SuperAddinCS4 OnClickRibbonButton");
            }
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            Utils.Dialog.ShowError(exception, "Unexpected state in SuperAddinCS4 " + methodKind.ToString());
        }

        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An register error occurend in " + methodKind.ToString(), "SuperAddinCS4");
        }
    }
}