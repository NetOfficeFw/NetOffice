using System;
using System.Windows.Forms;
using ExampleBase;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace ExcelExamplesCS4
{
    /// <summary>
    /// Example 6 - Dialogs in Excel
    /// </summary>
    partial class Example06 : UserControl , IExample
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public Example06()
        {
            InitializeComponent();
        }

        #endregion

        #region IExample Member

        public void RunExample()
        {
            // its an example with an own visual control
            // checkout buttonStartExample_Click
        }

        public void Connect(IHost hostApplication)
        {
            HostApplication = hostApplication;
        }

        public string Caption
        {
            get { return "Example06"; }
        }

        public string Description
        {
            get { return "Dialogs in Excel"; }
        }

        public UserControl Panel
        {
            get { return this; }
        }
    
        #endregion

        #region Properties

        /// <summary>
        /// Current Example Host
        /// </summary>
        internal IHost HostApplication { get; private set; }

        #endregion

        #region UI Trigger

        private void buttonStartExample_Click(object sender, EventArgs e)
        {           
            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            excelApplication.Visible = true;

            // add a new workbook
            Excel.Workbook workBook = excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = (Excel.Worksheet)workBook.Worksheets[1];

            // show selected window and display user clicks ok or cancel
            bool returnValue = false;
            RadioButton radioSelectButton = GetSelectedRadioButton();
            switch (radioSelectButton.Text)
            {
                case "xlDialogAddinManager":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogAddinManager].Show();
                    break;

                case "xlDialogFont":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogFont].Show();
                    break;

                case "xlDialogEditColor":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogEditColor].Show();
                    break;

                case "xlDialogGallery3dBar":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogGallery3dBar].Show();
                    break;

                case "xlDialogSearch":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogSearch].Show();
                    break;

                case "xlDialogPrinterSetup":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogPrinterSetup].Show();
                    break;

                case "xlDialogFormatNumber":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogFormatNumber].Show();
                    break;

                case "xlDialogApplyStyle":

                    returnValue = excelApplication.Dialogs[XlBuiltInDialog.xlDialogApplyStyle].Show();
                    break;

                default:
                    throw (new ArgumentOutOfRangeException("Unkown dialog selected."));

            }

            string message = string.Format("The dialog returns {0}.", returnValue);
            MessageBox.Show(this, message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();
        }
        
        #endregion

        #region Helper

        private RadioButton GetSelectedRadioButton()
        {
            foreach (Control itemControl in panelSelection.Controls)
            {
                RadioButton radioSelectButton = itemControl as RadioButton;
                if (null != radioSelectButton)
                {
                    if (radioSelectButton.Checked)
                        return radioSelectButton;
                }
            }

            throw (new InvalidOperationException("No Dialog selected."));
        }

        #endregion
    }
}
