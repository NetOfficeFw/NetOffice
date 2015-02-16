using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using ExampleBase;

using NetOffice;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using NetOffice.OfficeApi.Enums;
using VB = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

namespace PowerPointExamplesCS4
{
    /// <summary>
    /// Example 3 - Create Macros 
    /// </summary>
    internal  class Example03 : IExample
    {
        #region IExample

        public void RunExample()
        {
            bool isFailed = false;
            string documentFile = null;
            PowerPoint.Application powerApplication = null;
            try
            {
                // start powerpoint
                powerApplication = new PowerPoint.Application();

                // add a new presentation with one new slide
                PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
                PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

                // add new module and insert macro. the option "Trust access to Visual Basic Project" must be set
                VB.CodeModule module = presentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule;
                string macro = string.Format("Sub NetOfficeTestMacro()\r\n   {0}\r\nEnd Sub", "MsgBox \"Thanks for click!\"");
                module.InsertLines(1, macro);

                // add button and connect with macro
                PowerPoint.Shape button = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeActionButtonForwardorNext, 100, 100, 200, 200);
                button.ActionSettings[PpMouseActivation.ppMouseClick].AnimateAction = MsoTriState.msoTrue;
                button.ActionSettings[PpMouseActivation.ppMouseClick].Action = PpActionType.ppActionRunMacro;
                button.ActionSettings[PpMouseActivation.ppMouseClick].Run = "NetOfficeTestMacro";
               
                // save the document 
                string fileExtension = GetDefaultExtension(powerApplication);
                documentFile = string.Format("{0}\\Example03{1}", HostApplication.RootDirectory, fileExtension);
                presentation.SaveAs(documentFile);
            }
            catch (System.Runtime.InteropServices.COMException throwedException)
            {
                isFailed = true;
                HostApplication.ShowErrorDialog("VBA Error", throwedException);
            }
            finally
            {
                // close power point and dispose reference
                if (powerApplication != null)
                {
                    powerApplication.Quit();
                    powerApplication.Dispose();
                }

                if ((null != documentFile) && (!isFailed))
                    HostApplication.ShowFinishDialog(null, documentFile);
            }
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return HostApplication.LCID == 1033 ? "Example03" : "Beispiel03"; }
        }

        public string Description
        {
            get { return HostApplication.LCID == 1033 ? "Create an run macros. the option 'Trust access to Visual Basic Project' must be set" : "Makros erstellen und ausführen. Die Option 'Visual Basic Projekten vertrauen' muss aktiviert sein."; }
        }

        public UserControl Panel
        {
            get { return null; }
        }

        #endregion

        #region Properties

        internal IHost HostApplication { get; private set; }

        #endregion

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".ppt" or ".pptm"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(PowerPoint.Application application)
        {
            double Version = Convert.ToDouble(application.Version, CultureInfo.InvariantCulture);
            if (Version >= 12.00)
                return ".pptm";
            else
                return ".ppt";
        }

        #endregion
    }
}
