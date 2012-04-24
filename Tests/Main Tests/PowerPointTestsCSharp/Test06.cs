using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Drawing;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;

namespace PowerPointTestsCSharp
{
    public class Test06 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test06"; }
        }

        public string Description
        {
            get { return "Create custom UI."; }
        }

        public string OfficeProduct
        {
            get { return "PowerPoint"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            PowerPoint.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                Bitmap iconBitmap = new Bitmap(System.Reflection.Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("PowerPointTestsCSharp.Test06.bmp"));
                application = new PowerPoint.Application();

                Office.CommandBar commandBar;
                Office.CommandBarButton commandBarBtn;
  
                // add a new presentation with one new slide
                PowerPoint.Presentation presentation = application.Presentations.Add(MsoTriState.msoTrue);
                PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

                // add a commandbar popup
                Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarPopup.Caption = "commandBarPopup";
 
                #region CommandBarButton

                // add a button to the popup
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                Clipboard.SetDataObject(iconBitmap);
                commandBarBtn.PasteFace();
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                #region Create a new toolbar

                // add a new toolbar
                commandBar = application.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
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
                commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Frames"].Controls.Add(
                                                                MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
                commandBarPopup.Caption = "commandBarPopup";

                // add a button to the popup
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Missing.Value, Missing.Value, Missing.Value, true);
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

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Ctrl.Dispose();
        }

        #endregion
    }
}
