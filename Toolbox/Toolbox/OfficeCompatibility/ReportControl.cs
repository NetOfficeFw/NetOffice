using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.OfficeCompatibility
{
    public partial class ReportControl : UserControl
    {
        #region Fields

        AnalyzerResult _report;
        int _currentLanguageID;

        #endregion

        #region Construction

        public ReportControl(AnalyzerResult report, int currentLanguageID)
        {
            InitializeComponent();
            if (null == report.Report)
                return;
            _report = report;
            _currentLanguageID = currentLanguageID;
            comboBoxFilter.SelectedIndex = 0;
            Translator.TranslateControls(this, "OfficeCompatibility.ReportMessageTable.txt", _currentLanguageID);
        }
        
        #endregion

        #region Methods

        private void ShowAssembly()
        {
            treeViewReport.Nodes.Clear();
            foreach (XElement item in _report.Report.Document.Root.Elements("Assembly"))
            {
                string name = item.Attribute("Name").Value;
                if (name.IndexOf(",", StringComparison.InvariantCultureIgnoreCase) > 0)
                    name = name.Substring(0, name.IndexOf(",", StringComparison.InvariantCultureIgnoreCase));
                TreeNode node = treeViewReport.Nodes.Add(name);
                node.ImageIndex = 0;
                node.SelectedImageIndex = 0;
                node.Tag = item;

                foreach (XElement itemClass in item.Element("Classes").Elements("Class"))
                {
                    TreeNode classNode = node.Nodes.Add(itemClass.Attribute("Name").Value);
                    classNode.ImageIndex = 1;
                    classNode.SelectedImageIndex = 1;
                    classNode.Tag = itemClass;


                    foreach (XElement itemField in itemClass.Element("Fields").Elements("Entity"))
                    {
                        if (FilterPassed(itemField.Element("SupportByLibrary")))
                        { 
                            TreeNode fieldNode = classNode.Nodes.Add(itemField.Attribute("Name").Value);
                            fieldNode.ImageIndex = 2;
                            fieldNode.SelectedImageIndex = 2;
                            fieldNode.Tag = itemField;
                        }
                    }

                    foreach (XElement itemMethod in itemClass.Element("Methods").Elements("Method"))
                    {
                        bool filterPassed = false;

                        foreach (XElement itemFilterNode in itemMethod.Descendants("SupportByLibrary"))
                        {
                            if (FilterPassed(itemFilterNode))
                            {
                                filterPassed = true;
                                break;
                            }
                        }

                        if (filterPassed)
                        { 
                            TreeNode methodNode = classNode.Nodes.Add(itemMethod.Attribute("Name").Value);
                            methodNode.ImageIndex = 3;
                            methodNode.SelectedImageIndex = 3;
                            methodNode.Tag = itemMethod;
                        }
                    }

                }
            }
            panelView.Dock = DockStyle.Fill;
            panelNativeView.Dock = DockStyle.Fill;

            if (treeViewReport.Nodes.Count > 0)
                treeViewReport.Nodes[0].Expand();
        }

        private void SetMethodCalls(XElement element)
        {
            XElement parametersNode = element.Element("Calls");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                ListViewItem viewItem = listView1.Items.Add("Calls");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);

                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    string name = item.Attribute("Name").Value;
                    if (name.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) > -1)
                        name = name.Substring(0,name.IndexOf("::", StringComparison.InvariantCultureIgnoreCase));

                    string type = item.Element("SupportByLibrary").Attribute("Name").Value;
                    if (type.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) > -1)
                        type = type.Substring(type.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) + 2);

                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        return;

                    ListViewItem paramViewItem = listView1.Items.Add(name);
                    paramViewItem.SubItems.Add(type);
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem versionViewItem = listView1.Items.Add("");
                        versionViewItem.SubItems.Add("");
                        versionViewItem.SubItems.Add(itemVersion.Value);
                    }

                    XElement itemParameters = item.Element("Parameters");
                    if (null != itemParameters)
                    {
                        foreach (XElement paramItem in itemParameters.Elements("Parameter"))
	                    {
                            if (!FilterPassed(paramItem.Element("SupportByLibrary")))
                                return;

                            ListViewItem paramViewItem2 = listView1.Items.Add("Parameter");
                            paramViewItem2.SubItems.Add(paramItem.Element("SupportByLibrary").Attribute("Name").Value);
                            foreach (XElement itemVersionParam in paramItem.Element("SupportByLibrary").Elements("Version"))
                            {
                                  ListViewItem paramViewVersionItem2 = listView1.Items.Add("");
                                  paramViewVersionItem2.SubItems.Add("");
                                  paramViewVersionItem2.SubItems.Add(itemVersionParam.Value);
                            }
	                    }
                        
                    }
                }
            }
        }

        private void SetMethodLocalFieldSets(XElement element)
        {
            XElement parametersNode = element.Element("LocalFieldSets");
            if ((null != parametersNode) && (parametersNode.Elements("Field").Count() > 0))
            {
                ListViewItem viewItem = listView1.Items.Add("LocalFieldSets");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);
                foreach (XElement item in parametersNode.Elements("Field"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        return;

                    ListViewItem paramViewItem = listView1.Items.Add(item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Element("SupportByLibrary").Attribute("Name").Value);
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem versionViewItem = listView1.Items.Add("");
                        versionViewItem.SubItems.Add("");
                        versionViewItem.SubItems.Add(itemVersion.Value);
                    }
                }
            }
        }

        private void SetMethodFieldSets(XElement element)
        {
            XElement parametersNode = element.Element("FieldSets");
            if ((null != parametersNode) && (parametersNode.Elements("Field").Count() > 0))
            {
                ListViewItem viewItem = listView1.Items.Add("FieldSets");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);

                foreach (XElement item in parametersNode.Elements("Field"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        return;

                    ListViewItem paramViewItem = listView1.Items.Add(item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Element("SupportByLibrary").Attribute("Name").Value);
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem versionViewItem = listView1.Items.Add("");
                        versionViewItem.SubItems.Add("");
                        versionViewItem.SubItems.Add(itemVersion.Value);
                    }
                }
            }
        }

        private void SetMethodVariables(XElement element)
        {
            XElement parametersNode = element.Element("Variables");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                ListViewItem viewItem = listView1.Items.Add("Variables");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);

                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        return;

                    ListViewItem paramViewItem = listView1.Items.Add(item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Attribute("Type").Value);
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem versionViewItem = listView1.Items.Add("");
                        versionViewItem.SubItems.Add("");
                        versionViewItem.SubItems.Add(itemVersion.Value);
                    }
                }
            }
        }

        private void SetMethodParameters(XElement element)
        {
            XElement parametersNode = element.Element("Parameters");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                ListViewItem viewItem = listView1.Items.Add("Parameters");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);

                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        return;

                    ListViewItem paramViewItem = listView1.Items.Add(item.Attribute("Type").Value);
                    paramViewItem.SubItems.Add(item.Attribute("Name").Value);
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem versionViewItem = listView1.Items.Add("");
                        versionViewItem.SubItems.Add("");
                        versionViewItem.SubItems.Add(itemVersion.Value);
                    }
                }
            }
        }

        private void SetMethodReturnValue(XElement element)
        {
            XElement returnValueNode = element.Element("ReturnValue");
            if (null != returnValueNode)
            {
                if (!FilterPassed(returnValueNode.Element("Entity").Element("SupportByLibrary")))
                    return;

                string valType = returnValueNode.Element("Entity").Attribute("Type").Value;
                ListViewItem viewItem = listView1.Items.Add("Return Value");
                viewItem.Font = new Font(viewItem.Font, FontStyle.Bold);

                viewItem.SubItems.Add(valType);
                foreach (XElement versionItem in returnValueNode.Element("Entity").Element("SupportByLibrary").Elements("Version"))
                {
                    ListViewItem versionViewItem = listView1.Items.Add("");
                    versionViewItem.SubItems.Add("");
                    versionViewItem.SubItems.Add(versionItem.Value);
                }

            }
        }

        private bool FilterPassed(XElement supportNode)
        {
            if (0 == comboBoxFilter.SelectedIndex)
                return true;

            bool found09 = false;
            bool found10 = false;
            bool found11 = false;
            bool found12 = false;
            bool found14 = false;

            foreach (XElement itemVersion in supportNode.Elements("Version"))
            {
                switch (itemVersion.Value)
                {
                    case "9":
                        found09 = true;
                        break;
                    case "10":
                        found10 = true;
                        break;
                    case "11":
                        found11 = true;
                        break;
                    case "12":
                        found12 = true;
                        break;
                    case "14":
                        found14 = true;
                        break;
                    default:
                        break;
                }
            }

            switch (comboBoxFilter.SelectedIndex)
            {
                case 1:     // 09
                    if (found09)
                        return false;
                    break;
                case 2:     // 10
                    if (found10)
                        return false;
                    break;
                case 3:     // 11
                    if (found11)
                        return false;
                    break;
                case 4:     // 12
                    if (found12)
                        return false;
                    break;
                case 5:     // 14
                    if (found14)
                        return false;
                    break;
            }

            return true;
        }

        #endregion

        #region Trigger

        private void buttonClose2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void treeViewReport_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (null == treeViewReport.SelectedNode)
            {
                textBoxReport.Text = "";
                return;
            }

            XElement element = treeViewReport.SelectedNode.Tag as XElement;
            switch (element.Name.ToString())
            {
                case "Assembly":
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns[0].Width = 200;
                    listView1.Items.Add(element.Attribute("Name").Value);
                    break;
                case "Class":
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns[0].Width = 200;
                    listView1.Items.Add(element.Attribute("Name").Value);
                    break;
                case "Entity":
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns.Add("");
                    listView1.Columns.Add("Support");
                    listView1.Columns[0].Width = 200;
                    listView1.Columns[1].Width = 200;
                    if (!FilterPassed(element.Element("SupportByLibrary")))
                        break;
                    listView1.Items.Add(element.Attribute("Name").Value);
                    listView1.Items[0].SubItems.Add(element.Attribute("Type").Value);
                    foreach (XElement item in element.Element("SupportByLibrary").Elements("Version"))
                    {
                        ListViewItem viewItem = listView1.Items.Add("");
                        viewItem.SubItems.Add("");
                        viewItem.SubItems.Add(item.Value);
                    }
                    break;
                case "Method":
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns.Add("");
                    listView1.Columns.Add("Support");
                    listView1.Columns[0].Width = 200;
                    listView1.Columns[1].Width = 200;
                    SetMethodReturnValue(element);
                    SetMethodParameters(element);
                    SetMethodVariables(element);
                    SetMethodLocalFieldSets(element);
                    SetMethodFieldSets(element);
                    SetMethodCalls(element);
                    break;
                default:
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    break;
            }

            textBoxReport.Text = element.ToString();
        }

        private void checkBoxNativeView_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                panelView.Visible = !checkBoxNativeView.Checked;
                panelNativeView.Visible = checkBoxNativeView.Checked;
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
            
        }

        private void comboBoxFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ShowAssembly();
            }
            catch (Exception exception)
            {
                ErrorForm errorForm = new ErrorForm(exception, ErrorCategory.NonCritical, _currentLanguageID);
                errorForm.ShowDialog(this);
            }
           
        }

        #endregion
    }
}
