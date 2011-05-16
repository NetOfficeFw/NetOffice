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
using NetOffice.OfficeApi.Enums; 

namespace Example04
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Initialize Api COMObject Support
            LateBindingApi.Core.Factory.Initialize();

            // start word and turn off msg boxes
            PowerPoint.Application powerApplication = new PowerPoint.Application();
            powerApplication.DisplayAlerts = PpAlertLevel.ppAlertsNone;

            // add a new presentation with two new slides
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide1 = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);
            PowerPoint.Slide slide2 = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add shapes
            slide1.Shapes.AddShape(MsoAutoShapeType.msoShape4pointStar, 100, 100, 200, 200);
            slide2.Shapes.AddShape(MsoAutoShapeType.msoShapeDoubleWave, 200, 200, 200, 200);

            // change blend animation
            slide1.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverDown;
            slide1.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast;
            
            slide2.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectCoverLeftDown;
            slide2.SlideShowTransition.Speed = PpTransitionSpeed.ppTransitionSpeedFast;

            // save the document 
            string fileExtension = GetDefaultExtension(powerApplication);
            string documentFile = string.Format("{0}\\Example04{1}", Environment.CurrentDirectory, fileExtension);
            presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

            // close power point and dispose reference
            powerApplication.Quit();
            powerApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Presentation saved.", (string)documentFile);
            fDialog.ShowDialog(this);
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
