using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeCompatibility
{
    [RessourceTable("ToolboxControls.OfficeCompatibility.Report.txt")]
    public partial class ReportControl : UserControl, ILocalizationDesign
    {
        #region Fields

        AnalyzerResult _report;
        int _currentLanguageID;

        #endregion

        #region Construction

        public ReportControl()
        {
            InitializeComponent();

            pictureBoxField.Image = imageList1.Images[3];
            pictureBoxProperty.Image = imageList1.Images[7];
            pictureBoxMethod.Image = imageList1.Images[5];
        }

        public ReportControl(AnalyzerResult report, int currentLanguageID)
        {
            InitializeComponent();
            if (null == report.Report)
                return;
            _report = report;
            _currentLanguageID = currentLanguageID;
            comboBoxFilter.SelectedIndex = 0;

            Translation.Translator.AutoTranslateControls(this, "OfficeCompatibility - Report", "ToolboxControls.OfficeCompatibility.Report.txt", currentLanguageID);

            pictureBoxField.Image = imageList1.Images[3];
            pictureBoxProperty.Image = imageList1.Images[7];
            pictureBoxMethod.Image = imageList1.Images[5];
        }
        
        #endregion

        #region Methods

        private int GetPercent(int value, int percent)
        {
            return value / 100 * percent;
        }

        private int GetImageClassIndex(XElement itemClass)
        {
            if (itemClass.Attribute("IsPublic").Value.Equals("true", StringComparison.InvariantCultureIgnoreCase))
                return 1;
            else
                return 2;
        }

        private int GetImageFieldIndex(XElement itemField)
        {
            if (itemField.Attribute("IsPublic").Value.Equals("true", StringComparison.InvariantCultureIgnoreCase))
                return 3;
            else
                return 4;
        }

        private int GetImageMethodIndex(XElement itemMethod)
        {
            if (itemMethod.Attribute("IsProperty").Value.Equals("true", StringComparison.InvariantCultureIgnoreCase))
            {
                if (itemMethod.Attribute("IsPublic").Value.Equals("true", StringComparison.InvariantCultureIgnoreCase))
                    return 7;
                else
                    return 8;
            }
            else
            {
                if (itemMethod.Attribute("IsPublic").Value.Equals("true", StringComparison.InvariantCultureIgnoreCase))
                    return 5;
                else
                    return 6;
            }
        }

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
                    classNode.ImageIndex = GetImageClassIndex(itemClass);
                    classNode.SelectedImageIndex = GetImageClassIndex(itemClass);
                    classNode.Tag = itemClass;


                    foreach (XElement itemField in itemClass.Element("Fields").Elements("Entity"))
                    {
                        if (FilterPassed(itemField.Element("SupportByLibrary")))
                        { 
                            TreeNode fieldNode = classNode.Nodes.Add(itemField.Attribute("Name").Value);
                            fieldNode.ImageIndex = GetImageFieldIndex(itemField);
                            fieldNode.SelectedImageIndex = GetImageFieldIndex(itemField);
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
                            string methodName = itemMethod.Attribute("Name").Value;
                            if(methodName.StartsWith("get_"))
                            {
                                methodName = methodName.Substring(4);
                                TreeNode methodNode = classNode.Nodes.Add(methodName, methodName);
                                methodNode.ImageIndex = GetImageMethodIndex(itemMethod);
                                methodNode.SelectedImageIndex = GetImageMethodIndex(itemMethod);
                                methodNode.Tag = itemMethod;
                            }
                            else if( methodName.StartsWith("set_"))
                            {
                                methodName = methodName.Substring(4);
                                List<XElement> list = new List<XElement>();
                                TreeNode getNode = classNode.Nodes[methodName];
                                list.Add(getNode.Tag as XElement);
                                list.Add(itemMethod);
                                getNode.Tag = list;
                            }
                            else
                            {
                                TreeNode methodNode = classNode.Nodes.Add(methodName);
                                methodNode.ImageIndex = GetImageMethodIndex(itemMethod);
                                methodNode.SelectedImageIndex = GetImageMethodIndex(itemMethod);
                                methodNode.Tag = itemMethod;
                            }
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
                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    string name = item.Attribute("Name").Value;
                    if (name.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) > -1)
                        name = name.Substring(0,name.IndexOf("::", StringComparison.InvariantCultureIgnoreCase));

                    string type = item.Element("SupportByLibrary").Attribute("Name").Value;
                    if (type.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) > -1)
                        type = type.Substring(type.IndexOf("::", StringComparison.InvariantCultureIgnoreCase) + 2);

                    if (FilterPassed(item.Element("SupportByLibrary")))
                    { 
                        ListViewItem paramViewItem = listView1.Items.Add("Call Method: " + name);
                        paramViewItem.SubItems.Add(type);
                        string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                        foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                            supportText += itemVersion.Value + " ";
                        paramViewItem.SubItems.Add(supportText);

                    }
                    XElement itemParameters = item.Element("Parameters");
                    if (null != itemParameters)
                    {
                        foreach (XElement paramItem in itemParameters.Elements("Parameter"))
	                    {
                            if (!FilterPassed(paramItem.Element("SupportByLibrary")))
                                continue;

                            ListViewItem paramViewItem2 = listView1.Items.Add("   Parameter");
                            paramViewItem2.SubItems.Add(paramItem.Element("SupportByLibrary").Attribute("Name").Value);
                            string supportText = paramItem.Element("SupportByLibrary").Attribute("Api").Value + " ";
                            foreach (XElement itemVersionParam in paramItem.Element("SupportByLibrary").Elements("Version"))
                                supportText += itemVersionParam.Value + " ";
                            paramViewItem2.SubItems.Add(supportText);

	                    }
                        
                    }
                }
                listView1.Items.Add("");
            }
        }

        private void SetMethodLocalFieldSets(XElement element)
        {
            XElement parametersNode = element.Element("LocalFieldSets");
            if ((null != parametersNode) && (parametersNode.Elements("Field").Count() > 0))
            {
                foreach (XElement item in parametersNode.Elements("Field"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        continue;

                    ListViewItem paramViewItem = listView1.Items.Add("Set Local Variable " + item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Element("SupportByLibrary").Attribute("Name").Value);
                    string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                        supportText += itemVersion.Value + " ";
                    paramViewItem.SubItems.Add(supportText);
                }
                listView1.Items.Add("");
            }
        }

        private void SetMethodFieldSets(XElement element)
        {
            XElement parametersNode = element.Element("FieldSets");
            if ((null != parametersNode) && (parametersNode.Elements("Field").Count() > 0))
            {
                foreach (XElement item in parametersNode.Elements("Field"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        continue;

                    ListViewItem paramViewItem = listView1.Items.Add("Set Class Field " + item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Element("SupportByLibrary").Attribute("Name").Value);
                    string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                        supportText += itemVersion.Value + " ";
                    paramViewItem.SubItems.Add(supportText);
                }
                listView1.Items.Add("");
            }
        }

        private void SetNewObjects(XElement element)
        {
            XElement parametersNode = element.Element("NewObjects");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        continue;

                    ListViewItem paramViewItem = listView1.Items.Add("new " + item.Attribute("Type").Value + "()");
                    paramViewItem.SubItems.Add(item.Attribute("Type").Value);
                    string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                        supportText += itemVersion.Value + " ";
                    paramViewItem.SubItems.Add(supportText);
                }
                listView1.Items.Add("");
            }
        }

        private void SetMethodVariables(XElement element)
        {
            XElement parametersNode = element.Element("Variables");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        continue;

                    ListViewItem paramViewItem = listView1.Items.Add("Locale Variable " + item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Attribute("Type").Value);
                    string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                        supportText += itemVersion.Value + " ";
                    paramViewItem.SubItems.Add(supportText);
                }
                listView1.Items.Add("");
            }
        }

        private void SetMethodParameters(XElement element)
        {
            XElement parametersNode = element.Element("Parameters");
            if ((null != parametersNode) && (parametersNode.Elements("Entity").Count() > 0))
            {
                foreach (XElement item in parametersNode.Elements("Entity"))
                {
                    if (!FilterPassed(item.Element("SupportByLibrary")))
                        continue;

                    ListViewItem paramViewItem = listView1.Items.Add("Parameter " + item.Attribute("Name").Value);
                    paramViewItem.SubItems.Add(item.Attribute("Type").Value);

                    string supportText = item.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement itemVersion in item.Element("SupportByLibrary").Elements("Version"))
                        supportText += itemVersion.Value + " ";
                    paramViewItem.SubItems.Add(supportText);
                }
                listView1.Items.Add("");
            }
        }

        private void SetMethodReturnValue(XElement element)
        {
            XElement returnValueNode = element.Element("ReturnValue");
            if (null != returnValueNode)
            {
                if (!FilterPassed(returnValueNode.Element("Entity").Element("SupportByLibrary")))
                    return;

                string valType = returnValueNode.Element("Entity").Attribute("FullType").Value;
                ListViewItem viewItem = listView1.Items.Add("Return Value");

                viewItem.SubItems.Add(valType);
                string supportText = returnValueNode.Element("Entity").Element("SupportByLibrary").Attribute("Api").Value + " ";
                foreach (XElement versionItem in returnValueNode.Element("Entity").Element("SupportByLibrary").Elements("Version"))
                    supportText += versionItem.Value + " ";
                viewItem.SubItems.Add(supportText);

                listView1.Items.Add("");
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
            bool found15 = false;

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
                    case "15":
                        found15 = true;
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
                case 6:     // 15
                    if (found15)
                        return false;
                    break;
            }

            return true;
        }

        private string GetLogFileContent()
        {
            return _report.Report.ToString();
        }

        #endregion

        #region ILocalizationDesign

        public void EnableDesignView(int lcid, string parentComponentName)
        {
           
        }

        public void Localize(Translation.ItemCollection strings)
        {
            Translation.Translator.TranslateControls(this, strings);
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
        }

        public IContainer Components
        {
            get { return components; }
        }

        public string NameLocalization
        {
            get
            {
                return null;
            }
        }

        public IEnumerable<ILocalizationChildInfo> Childs
        {
            get
            {
                return new ILocalizationChildInfo[0];
            }
        }

        #endregion

        #region Trigger

        private void buttonClose2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void ShowDetails(List<XElement> elements)
        {
            listView1.Items.Clear();
            foreach (XElement element in elements)
                ShowDetails(element, false);
        }

        private void ShowDetails(XElement element, bool clearOldItems)
        {
            switch (element.Name.ToString())
            {
                case "Assembly":
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns[0].Width = listView1.Width - 50;
                    listView1.Columns[0].Text = "Assenbly";
                    listView1.Items.Add(element.Attribute("Name").Value);
                    break;
                case "Class":
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns[0].Width = listView1.Width - 50;
                    listView1.Columns[0].Text = "Class";
                    listView1.Items.Add(element.Attribute("Name").Value);
                    break;
                case "Entity":

                    listView1.Items.Clear();
                    listView1.Columns.Clear();

                    listView1.Columns.Add("Name");
                    listView1.Columns.Add("Type");
                    listView1.Columns.Add("Support");

                    listView1.Columns[0].Width = GetPercent(listView1.Width, 25);
                    listView1.Columns[0].Tag = 25;
                    listView1.Columns[1].Width = GetPercent(listView1.Width, 50);
                    listView1.Columns[1].Tag = 50;
                    listView1.Columns[2].Width = GetPercent(listView1.Width, 25);
                    listView1.Columns[2].Tag = 25;

                    if (!FilterPassed(element.Element("SupportByLibrary")))
                        break;

                    listView1.Items.Add(element.Attribute("Name").Value);
                    listView1.Items[0].SubItems.Add(element.Attribute("Type").Value);

                    string supportText = element.Element("SupportByLibrary").Attribute("Api").Value + " ";
                    foreach (XElement item in element.Element("SupportByLibrary").Elements("Version"))
                        supportText += item.Value + " ";

                    listView1.Items[0].SubItems.Add(supportText);

                    break;
                case "Method":
                    if (clearOldItems)
                        listView1.Items.Clear();
                    listView1.Columns.Clear();
                    listView1.Columns.Add("");
                    listView1.Columns.Add("");
                    listView1.Columns.Add("Support");

                    listView1.Columns[0].Width = GetPercent(listView1.Width, 25);
                    listView1.Columns[0].Tag = 25;
                    listView1.Columns[1].Width = GetPercent(listView1.Width, 50);
                    listView1.Columns[1].Tag = 50;
                    listView1.Columns[2].Width = GetPercent(listView1.Width, 25);
                    listView1.Columns[2].Tag = 25;

                    SetMethodReturnValue(element);
                    SetMethodParameters(element);
                    SetMethodVariables(element);
                    SetMethodLocalFieldSets(element);
                    SetMethodFieldSets(element);
                    SetNewObjects(element);
                    SetMethodCalls(element);
                    break;
                default:
                    listView1.Items.Clear();
                    listView1.Columns.Clear();
                    break;
            }

            textBoxReport.Text = element.ToString();
        }

        private void treeViewReport_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (null == treeViewReport.SelectedNode)
            {
                textBoxReport.Text = "";
                return;
            }
            
            XElement element = treeViewReport.SelectedNode.Tag as XElement;
            if (null != element)
            {
                ShowDetails(element, true);
            }
            else
            {
                List<XElement> elements = treeViewReport.SelectedNode.Tag as List<XElement>;
                ShowDetails(elements);            
            }
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
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, _currentLanguageID);
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
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, _currentLanguageID);
            }
        }

        private void listView1_Resize(object sender, EventArgs e)
        {
            foreach (ColumnHeader item in listView1.Columns)
            {
                if (item.Tag != null)
                {
                    int percentValue = Convert.ToInt32(item.Tag);
                    item.Width = GetPercent(listView1.Width, percentValue);
                }
            }
        }

        private void buttonSaveReport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "*.txt|*.txt";
                if(DialogResult.OK == dialog.ShowDialog(this))
                {
                    if (File.Exists(dialog.FileName))
                        File.Delete(dialog.FileName);

                    string logFileContent = GetLogFileContent();
                    File.AppendAllText(dialog.FileName, logFileContent);
                }
            }
            catch (Exception exception)
            {
                Forms.ErrorForm.ShowError(exception,ErrorCategory.NonCritical, _currentLanguageID);
            }
        }

        #endregion

    }
}
