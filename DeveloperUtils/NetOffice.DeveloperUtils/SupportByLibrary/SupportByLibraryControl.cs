using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Linq;
using System.Xml.Linq;

using Mono.Cecil;
using Mono.Cecil.Cil;

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

            if (anyTag is string[])
            {
                string[] paramArray = anyTag as string[];
                paramArray[0] = paramArray[0].Trim().Replace("-", "").Replace("/", "");
                if(!System.IO.File.Exists(paramArray[0]))
                {
                    Console.WriteLine("NetOffice.DeveloperUtils");
                    Console.WriteLine("File not found " + paramArray[0]);
                    return;
                }

                AssemblyAnalyzerSettings settings = CreateSettings(paramArray);
                if (null != settings)
                {
                    Console.WriteLine("NetOffice.DeveloperUtils");
                    Console.WriteLine("Can't proceed Arguments");
                    return;
                }

                AssemblyDefinition assemblyDefinition = AssemblyDefinition.ReadAssembly(paramArray[0]);
                string result = AssemblyAnalyzer.AnalyzeAssembly(assemblyDefinition, settings);
                Console.Write(result);
            }
            else
            {
                dataGridView.Rows.Add(5);

                dataGridView.Rows[0].Cells[0].Value = "Excel";
                dataGridView.Rows[1].Cells[0].Value = "Word";
                dataGridView.Rows[2].Cells[0].Value = "Outlook";
                dataGridView.Rows[3].Cells[0].Value = "PowerPoint";
                dataGridView.Rows[4].Cells[0].Value = "Access";
                for (int i = 1; i <= 5; i++)
                {
                    dataGridView.Rows[0].Cells[i].Value = true;
                    dataGridView.Rows[1].Cells[i].Value = true;
                    dataGridView.Rows[2].Cells[i].Value = true;
                    dataGridView.Rows[3].Cells[i].Value = true;
                    dataGridView.Rows[4].Cells[i].Value = true;
                }

                dataGridView.BorderStyle = BorderStyle.None;
            }
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
            if (e.RowIndex < 0)
                return;

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

        private void buttonSelectAssembly_Click(object sender, EventArgs e)
        {
            if (!OneOrMoreLibraryVersionIsSelected())
            {
                MessageBox.Show(this, "Please select one or more Type Library Version first.", "Please select", MessageBoxButtons.OK, MessageBoxIcon.Error);   
                return;
            }

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "*.exe|*.exe|*.dll|*.dll|All Files|*.*";
            if (DialogResult.OK == dialog.ShowDialog(this))
            {
                AssemblyDefinition assemblyDefinition = AssemblyDefinition.ReadAssembly(dialog.FileName);
                textBoxAssembly.Text = assemblyDefinition.Name.ToString();
                string result = AssemblyAnalyzer.AnalyzeAssembly(assemblyDefinition, CreateSettings());               
                textBoxConsole.Text = result;
            }
        }

        private void SetSettings(AssemblyAnalyzerSettingsLibrary settings, string[] array)
        {
            for (int i = 1; i < array.Length; i++)
            {
                array[i] = array[i].Trim();
                switch (array[i])
                { 
                    case "9":
                    case "09":
                        settings.Version9 = true;    
                        break;
                    case "10":
                        settings.Version10 = true;
                    break;
                    case "11":
                        settings.Version11 = true;
                        break;
                    case "12":
                        settings.Version12 = true;
                        break;
                    case "14":
                        settings.Version14 = true;
                        break;
                }
            }
        }

        private AssemblyAnalyzerSettings CreateSettings(string[] paramArray)
        {
            try
            {
                AssemblyAnalyzerSettings settings = new AssemblyAnalyzerSettings();

                for (int i = 1; i < paramArray.Length; i++)
                {
                    string[] splitArray = paramArray[i].Replace("-","").Replace("/","").Trim().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    splitArray[0] = splitArray[0].ToLower();
                    switch (splitArray[0]) 
                    {
                        case "excel":
                            SetSettings(settings.Excel, splitArray);
                            break;
                        case "word":
                            SetSettings(settings.Word, splitArray);
                            break;
                        case "outlook":
                            SetSettings(settings.Outlook, splitArray);
                            break;
                        case "powerpoint":
                            SetSettings(settings.PowerPoint, splitArray);
                            break;
                        case "access":
                            SetSettings(settings.Access, splitArray);
                            break;
                        case "office":
                            SetSettings(settings.Office, splitArray);
                            break;
                    }
                }

                return settings;
            }
            catch
            {
                return null;
            }
            
        }

        private AssemblyAnalyzerSettings CreateSettings()
        {
            AssemblyAnalyzerSettings settings = new AssemblyAnalyzerSettings();
           
            settings.Excel.Version9 = (bool)dataGridView.Rows[0].Cells[1].Value;
            settings.Excel.Version10 = (bool)dataGridView.Rows[0].Cells[2].Value;
            settings.Excel.Version11 = (bool)dataGridView.Rows[0].Cells[3].Value;
            settings.Excel.Version12 = (bool)dataGridView.Rows[0].Cells[4].Value;
            settings.Excel.Version14 = (bool)dataGridView.Rows[0].Cells[5].Value;

            settings.Word.Version9 = (bool)dataGridView.Rows[1].Cells[1].Value;
            settings.Word.Version10 = (bool)dataGridView.Rows[1].Cells[2].Value;
            settings.Word.Version11 = (bool)dataGridView.Rows[1].Cells[3].Value;
            settings.Word.Version12 = (bool)dataGridView.Rows[1].Cells[4].Value;
            settings.Word.Version14 = (bool)dataGridView.Rows[1].Cells[5].Value;

            settings.Outlook.Version9 = (bool)dataGridView.Rows[2].Cells[1].Value;
            settings.Outlook.Version10 = (bool)dataGridView.Rows[2].Cells[2].Value;
            settings.Outlook.Version11 = (bool)dataGridView.Rows[2].Cells[3].Value;
            settings.Outlook.Version12 = (bool)dataGridView.Rows[2].Cells[4].Value;
            settings.Outlook.Version14 = (bool)dataGridView.Rows[2].Cells[5].Value;

            settings.PowerPoint.Version9 = (bool)dataGridView.Rows[3].Cells[1].Value;
            settings.PowerPoint.Version10 = (bool)dataGridView.Rows[3].Cells[2].Value;
            settings.PowerPoint.Version11 = (bool)dataGridView.Rows[3].Cells[3].Value;
            settings.PowerPoint.Version12 = (bool)dataGridView.Rows[3].Cells[4].Value;
            settings.PowerPoint.Version14 = (bool)dataGridView.Rows[3].Cells[5].Value;

            settings.Access.Version9 = (bool)dataGridView.Rows[4].Cells[1].Value;
            settings.Access.Version10 = (bool)dataGridView.Rows[4].Cells[2].Value;
            settings.Access.Version11 = (bool)dataGridView.Rows[4].Cells[3].Value;
            settings.Access.Version12 = (bool)dataGridView.Rows[4].Cells[4].Value;
            settings.Access.Version14 = (bool)dataGridView.Rows[4].Cells[5].Value;

            for (int i = 1; i <= 5; i++)
            {
                for (int y = 0; y <= 4; y++)
                {
                    bool  value = (bool)dataGridView.Rows[y].Cells[i].Value;
                    if (value)
                        settings.Office.Version9 = true;
                }                
            }

            return settings;
        }

        private bool OneOrMoreLibraryVersionIsSelected()
        {
            for (int i = 1; i <= 5; i++)
            {
                DataGridViewCheckBoxCell checkCell1 = dataGridView.Rows[0].Cells[i] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell checkCell2 = dataGridView.Rows[1].Cells[i] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell checkCell3 = dataGridView.Rows[2].Cells[i] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell checkCell4 = dataGridView.Rows[3].Cells[i] as DataGridViewCheckBoxCell;
                DataGridViewCheckBoxCell checkCell5 = dataGridView.Rows[4].Cells[i] as DataGridViewCheckBoxCell;

                if( (null != checkCell1) && (null != checkCell1.Value))
                {
                    bool value = (bool)checkCell1.Value;
                    if (true == value)
                        return true;
                }

                if ((null != checkCell2) && (null != checkCell2.Value))
                {
                    checkCell2 = dataGridView.Rows[1].Cells[i] as DataGridViewCheckBoxCell;
                    bool value = (bool)checkCell2.Value;
                    if (true == value)
                        return true;
                }

                if ((null != checkCell3) && (null != checkCell3.Value))
                {
                    checkCell3 = dataGridView.Rows[2].Cells[i] as DataGridViewCheckBoxCell;
                    bool value = (bool)checkCell3.Value;
                    if (true == value)
                        return true;
                }

                if ((null != checkCell4) && (null != checkCell4.Value))
                {
                    checkCell4 = dataGridView.Rows[3].Cells[i] as DataGridViewCheckBoxCell;
                    bool value = (bool)checkCell4.Value;
                    if (true == value)
                        return true;
                }

                if ((null != checkCell5) && (null != checkCell5.Value))
                {
                    checkCell5 = dataGridView.Rows[4].Cells[i] as DataGridViewCheckBoxCell;
                    bool value = (bool)checkCell5.Value;
                    if (true == value)
                        return true;
                }
            }
            return false;
        }
    }
}
