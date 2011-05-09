using System;
using System.Reflection;
using System.Drawing;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace Example6
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

            // start excel and turn off msg boxes
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            
            // dont show dialogs with an invisible excel
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
                    throw (new Exception("Unkown dialog selected."));

            }

            string message = string.Format("The dialog returns {0}.", returnValue);
            MessageBox.Show(this, message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
  
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

        }
        
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

            throw (new Exception("No Dialog selected."));
        }

        #endregion
    }
}
