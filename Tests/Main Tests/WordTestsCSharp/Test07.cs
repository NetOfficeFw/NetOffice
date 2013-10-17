using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using Tests.Core;
using NetOffice;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;

namespace WordTestsCSharp
{
    public class Test07 : ITestPackage
    {
        #region TestPackage Member

        public string Name
        {
            get { return "Test07"; }
        }

        public string Description
        {
            get { return "Create custom UI."; }
        }

        public string OfficeProduct
        {
            get { return "Word"; }
        }

        public string Language
        {
            get { return "C#"; }
        }

        public TestResult DoTest()
        {
            Word.Application application = null;
            DateTime startTime = DateTime.Now;
            try
            {
                Bitmap iconBitmap = new Bitmap(System.Reflection.Assembly.GetAssembly(this.GetType()).GetManifestResourceStream("WordTestsCSharp.Test07.bmp"));
                application = new Word.Application();
                application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                application.Documents.Add();

                Office.CommandBar commandBar;
                Office.CommandBarButton commandBarBtn;

                Word.Template normalDotTemplate = GetNormalDotTemplate(application);
                application.CustomizationContext = normalDotTemplate;

                // add a commandbar popup
                Office.CommandBarPopup commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Menu Bar"].Controls.Add(MsoControlType.msoControlPopup, MsoBarPosition.msoBarTop, System.Type.Missing, 1, true);
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
                commandBarPopup = (Office.CommandBarPopup)application.CommandBars["Text"].Controls.Add(MsoControlType.msoControlPopup, MsoBarPosition.msoBarTop, false, 1, true);
                commandBarPopup.Caption = "commandBarPopup";

                // add a button to the popup
                commandBarBtn = (Office.CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, true);
                commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
                commandBarBtn.Caption = "commandBarButton";
                commandBarBtn.FaceId = 9;
                commandBarBtn.ClickEvent += new Office.CommandBarButton_ClickEventHandler(commandBarBtn_Click);

                #endregion

                normalDotTemplate.Saved = true;
                 
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
                    application.Quit(WdSaveOptions.wdDoNotSaveChanges);
                    application.Dispose();
                }
            }
        }

        /// <summary>
        /// returns normal.dot template (normal.dotm in modern word versions)
        /// </summary>
        private Word.Template GetNormalDotTemplate(Word.Application application)
        {
            foreach (Word.Template installedTemplate in application.Templates)
            {
                if (installedTemplate.Name.StartsWith("normal", StringComparison.InvariantCultureIgnoreCase))
                {
                    return installedTemplate;
                }
                installedTemplate.Dispose();
            }

            throw new IndexOutOfRangeException("Template not found.");
        }

        void commandBarBtn_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Ctrl.Dispose();
        }

        #endregion
    }
}
