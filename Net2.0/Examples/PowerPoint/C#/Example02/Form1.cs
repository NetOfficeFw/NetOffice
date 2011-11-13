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

namespace Example02
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

            // start powerpoint 
            PowerPoint.Application powerApplication = new PowerPoint.Application();

            // add a new presentation with one new slide
            PowerPoint.Presentation presentation = powerApplication.Presentations.Add(MsoTriState.msoTrue);
            PowerPoint.Slide slide = presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            // add a label
            PowerPoint.Shape label = slide.Shapes.AddLabel(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 600, 20);
            label.TextFrame.TextRange.Text = "This slide and created Shapes are created by NetOffice example.";

            // add a line
            slide.Shapes.AddLine(10, 80, 700, 80);

            // add a wordart
            slide.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect9, "This a WordArt", "Arial", 20, 
                                           MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 150);

            // add a star
            slide.Shapes.AddShape(MsoAutoShapeType.msoShape24pointStar, 200, 200, 250, 250);  

            // save the document 
            string fileExtension = GetDefaultExtension(powerApplication);
            string documentFile = string.Format("{0}\\Example02{1}", Application.StartupPath, fileExtension);
            presentation.SaveAs(documentFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

            // close power point and dispose reference
            powerApplication.Quit();
            powerApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Presentation saved.", documentFile);
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
