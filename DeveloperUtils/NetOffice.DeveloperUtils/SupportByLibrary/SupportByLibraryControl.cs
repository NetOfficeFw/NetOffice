using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperUtils.SupportByLibrary
{
    public partial class SupportByLibraryControl : UserControl, IUtilsControl
    {
        public SupportByLibraryControl()
        {
            InitializeComponent();
        }

        public SupportByLibraryControl(object anyTag)
        {
            InitializeComponent();

            dataGridView.Rows.Add(5);

            dataGridView.Rows[0].Cells[0].Value = "Excel";
            dataGridView.Rows[1].Cells[0].Value = "Word";
            dataGridView.Rows[2].Cells[0].Value = "Outlook";
            dataGridView.Rows[3].Cells[0].Value = "PowerPoint";
            dataGridView.Rows[4].Cells[0].Value = "Access";
            
            dataGridView.BorderStyle = BorderStyle.None;
        }

        #region IUtilsControl Members

        public string ControlName
        {
            get { return "SupportByLibrary"; }
        }

        public void Activate()
        {
           
        }

        public void LoadConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SaveConfiguration(System.Xml.XmlNode configNode)
        {
            
        }

        public void SetLanguage(int id)
        {
           
        }

        public void Release()
        {
           
        }

        #endregion

        private void dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex >= 6)
            {
                DataGridViewRow row = dataGridView.Rows[e.RowIndex];
                for(int i=1; i<=5; i++)
                {
                    DataGridViewCheckBoxCell checkCell = dataGridView.Rows[e.RowIndex].Cells[i] as DataGridViewCheckBoxCell;
                    if (6 == e.ColumnIndex)
                        checkCell.Value = true;
                    else
                        checkCell.Value = false;
                }
            }
        }
    }
}
