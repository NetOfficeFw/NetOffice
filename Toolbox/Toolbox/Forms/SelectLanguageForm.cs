using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Forms
{
    public partial class SelectLanguageForm : Form
    {
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect, // x-coordinate of upper-left corner
            int nTopRect, // y-coordinate of upper-left corner
            int nRightRect, // x-coordinate of lower-right corner
            int nBottomRect, // y-coordinate of lower-right corner
            int nWidthEllipse, // height of ellipse
            int nHeightEllipse // width of ellipse
         );

        public SelectLanguageForm()
        {
            InitializeComponent();
        }

        public SelectLanguageForm(string header)
        {
            InitializeComponent();

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = Forms.MainForm.Singleton.Languages;
            if (!String.IsNullOrWhiteSpace(header))
                labelHeader.Text = header;

            Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 10, 10));
        }

        public Translation.ToolLanguage Selected
        {
            get 
            {
                if (dataGridView1.SelectedCells.Count > 0)
                    return dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex].DataBoundItem as Translation.ToolLanguage;
                else
                    return null;
            }
        }

        public static Translation.ToolLanguage ShowForm(IWin32Window owner, string header = null)
        {
            SelectLanguageForm dlg = new SelectLanguageForm(header);
            DialogResult dr = dlg.ShowDialog(owner);
            if (dr == DialogResult.OK)
                return dlg.Selected;
            else
                return null;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            buttonSelect.Enabled = dataGridView1.SelectedCells.Count > 0;
        }

        private void buttonSelect_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedCells.Count > 0)
                buttonSelect_Click(buttonSelect, EventArgs.Empty);
        }
    }
}
