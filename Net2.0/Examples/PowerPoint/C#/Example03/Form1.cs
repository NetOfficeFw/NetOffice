using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using PowerPoint = NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using VBE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums; 
using NetOffice.OfficeApi.Enums; 

namespace Example03
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PowerPoint.Application powerApplication = null;
            try
            {
                // Initialize Api COMObject Support
                LateBindingApi.Core.Factory.Initialize();

                // start word and turn off msg boxes
                powerApplication = new PowerPoint.Application();
                powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone;

                // add a new presentation with one new slide
                PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
                PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

                // add new module and insert macro
                // the option "Trust access to Visual Basic Project" must be set
                VBE.CodeModule module = presentation.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule).CodeModule;
                string macro = string.Format("Sub NetOfficeTestMacro()\r\n   {0}\r\nEnd Sub", "MsgBox \"Thanks for click!\"");
                module.InsertLines(1, macro);

                // add button and connect with macro
                PowerPoint.Shape button = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeActionButtonForwardorNext, 100, 100, 200, 200);
                button.ActionSettings[PpMouseActivation.ppMouseClick].AnimateAction = MsoTriState.msoTrue;
                button.ActionSettings[PpMouseActivation.ppMouseClick].Action = PpActionType.ppActionRunMacro;
                button.ActionSettings[PpMouseActivation.ppMouseClick].Run = "NetOfficeTestMacro";
 
                // save the document 
                string fileExtension = GetDefaultExtension(powerApplication);
                string documentFile = string.Format("{0}\\Example03{1}", Environment.CurrentDirectory, fileExtension);
                presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

                 
                FinishDialog fDialog = new FinishDialog("Presentation saved.", (string)documentFile);
                fDialog.ShowDialog(this);
            }
            catch (Exception throwedException)
            {
                // not trusted
                string message = string.Format("An error is occured.{0}ExceptionTrace:{0}", Environment.NewLine);

                Exception exception = throwedException;
                while (null != exception)
                {
                    message += string.Format("{0}{1}", exception.Message, Environment.NewLine);
                    exception = exception.InnerException;
                }

                MessageBox.Show(message); 
            }
            finally
            { 
                // close power point and dispose reference
                if (powerApplication != null)
                {
                    powerApplication.Quit();
                    powerApplication.Dispose();
                    powerApplication = null;
                }
            }

        }

        #region Helper

        /// <summary>
        /// returns the valid file extension for the instance. for example ".ppt" or ".pptx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(PowerPoint.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".pptx";
            else
                return ".ppt";
        }

        #endregion
    }
}
