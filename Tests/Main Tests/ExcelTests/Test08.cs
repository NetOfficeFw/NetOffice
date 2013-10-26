using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using Core = NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Excel = NetOffice.ExcelApi;

namespace ExcelTestsCSharp
{
    /// <summary>
    /// foreach enumeration
    /// </summary>
    public class Test08 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test08"; }
        }

        public string Description
        {
            get { return "Create custom UI."; }
        }

        public string OfficeProduct
        {
            get { return "Excel"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Excel.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                Bitmap iconBitmap = new Bitmap(System.Reflection.Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("ExcelTestsCSharp.Test08.bmp"));
                application = new NetOffice.ExcelApi.Application();
                application.DisplayAlerts = false;
                application.Workbooks.Add();
                Excel.Worksheet sheet = application.Workbooks[1].Sheets[1] as Excel.Worksheet;

                // add a commandbar popup
                Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Worksheet Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarPopup.Caption = "commandBarPopup";

                #region CommandBarButton

                // add a button to the popup
                Office.CommandBarButton commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                Clipboard.SetDataObject(iconBitmap);
                commandBarBtn.PasteFace();
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                #region Create a new toolbar

                // add a new toolbar
                Office.CommandBar commandBar = application.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
                commandBar.Visible = true;

                // add a button to the toolbar
                commandBarBtn = (Office.CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                commandBarBtn.FaceId = 3;
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                // add a dropdown box to the toolbar
                commandBarPopup = (Office.CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarPopup.Caption = "commandBarPopup";

                // add a button to the popup, we use an own icon for the button
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                Clipboard.SetDataObject(iconBitmap);
                commandBarBtn.PasteFace();
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                #region Create a new ContextMenu

                // add a commandbar popup
                commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Cell"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarPopup.Caption = "commandBarPopup";

                // add a button to the popup
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                commandBarBtn.FaceId = 9;
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                return new TestResult(true, DateTime.Now.Subtract(startTime), "", null, "");
            }
            catch (Exception exception)
            {
                return new TestResult(false, DateTime.Now.Subtract(startTime), exception.Message, exception, "");
            }
            finally
            {
                if (null != application)
                {
                    application.Quit();
                    application.Dispose();
                }
            }
        }

        #endregion

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Ctrl.Dispose();
        }
    }
}
