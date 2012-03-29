using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace NetOffice.ProjectWizard
{
    public class AutomationAddinProject : NetOfficeProject
    {
        #region Fields

        string _registerCode;
        string _createClassicUIMethodCode;
        string _removeClassicUIMethodCode;
        string _createClassicUICallCode;
        string _removeClassicUICallCode;

        string _ribbonImplement;
        string _ribbonImplementCode;
        string _ribbonRessourceReference;
        string _ribbonHelperMethod;

        #endregion

        #region Properties

        public string RegisterCode
        {
            get
            {
                if (null == _registerCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _registerCode = ReadString("RegisterCodeCSharp.txt");
                    else
                        _registerCode = ReadString("RegisterCodeVisualBasic.txt");
                }
                return _registerCode;
            }
        }

        public string RibbonHelperMethod
        {
            get
            {
                if (null == _ribbonHelperMethod)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _ribbonHelperMethod = ReadString("ReadRessourceFileMethodCSharp.txt");
                    else
                        _ribbonHelperMethod = ReadString("ReadRessourceFileMethodVisualBasic.txt");
                }
                return _ribbonHelperMethod;
            }
        }

        public string RibbonRessourceReference
        {
            get
            {
                if (null == _ribbonRessourceReference)
                    _ribbonRessourceReference = ReadString("RibbonRessourceReference.txt");
                return _ribbonRessourceReference;
            }
        }

        public string RibbonImplement
        {
            get
            {
                if (null == _ribbonImplement)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _ribbonImplement = ReadString("RibbonImplementCSharp.txt");
                    else
                        _ribbonImplement = ReadString("RibbonImplementVisualBasic.txt");
                }
                return _ribbonImplement;
            }
        }

        public string RibbonImplementCode
        {
            get
            {
                if (null == _ribbonImplementCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _ribbonImplementCode = ReadString("RibbonImplementCodeCSharp.txt");
                    else
                        _ribbonImplementCode = ReadString("RibbonImplementCodeVisualBasic.txt");
                }
                return _ribbonImplementCode;
            }
        }

        public string CreateClassicUIMethodCode
        {
            get
            {
                if (null == _createClassicUIMethodCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _createClassicUIMethodCode = ReadString("CreateClassicUIMethodCSharp.txt");
                    else
                        _createClassicUIMethodCode = ReadString("CreateClassicUIMethodVisualBasic.txt");
                }
                return _createClassicUIMethodCode;
            }
        }

        public string RemoveClassicUIMethodCode
        {
            get
            {
                if (null == _removeClassicUIMethodCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _removeClassicUIMethodCode = ReadString("RemoveClassicUIMethodCSharp.txt");
                    else
                        _removeClassicUIMethodCode = ReadString("RemoveClassicUIMethodVisualBasic.txt");
                }
                return _removeClassicUIMethodCode;
            }
        }

        public string CreateClassicUICallCode
        {
            get
            {
                if (null == _createClassicUICallCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _createClassicUICallCode = ReadString("CreateUICallCodeCSharp.txt");
                    else
                        _createClassicUICallCode = ReadString("CreateUICallCodeVisualBasic.txt");
                }
                return _createClassicUICallCode;
            }
        }

        public string RemoveClassicUICallCode
        {
            get
            {
                if (null == _removeClassicUICallCode)
                {
                    if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                        _removeClassicUICallCode = ReadString("RemoveUICallCodeCSharp.txt");
                    else
                        _removeClassicUICallCode = ReadString("RemoveUICallCodeVisualBasic.txt");
                }
                return _removeClassicUICallCode;
            }
        }

        #endregion

        #region Methods
        
        string GetAddinHiveKey()
        {
            if ((ListControls[2] as IWizardControl).SettingsDocument.FirstChild.ChildNodes[0].InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
                return "LocalMachine";
            else
                return "CurrentUser";
        }

        string GetAddinOfficeKey(string officeApp)
        {
            string projectName = ReplacementsDictionary["$safeprojectname$"];
            string itemName = "Addin";
            return string.Format("Software\\Microsoft\\Office\\{0}\\Addins\\{1}.{2}", officeApp, projectName, itemName);
        }

        string GetAddinName()
        {
            string name = (ListControls[1] as IWizardControl).SettingsDocument.FirstChild.ChildNodes[0].InnerText;
            return name;
        }

        string GetAddinDescription()
        {
            string description = (ListControls[1] as IWizardControl).SettingsDocument.FirstChild.ChildNodes[1].InnerText;
            return description;
        }

        string GetAddinLoadBehavior()
        {
            return (ListControls[2] as IWizardControl).SettingsDocument.FirstChild.ChildNodes[1].InnerText;
        }

        void SetClassicUI()
        {
            IWizardControl uiControl = (ListControls[3] as IWizardControl);
            bool classicUIEnabled = uiControl.SettingsDocument.FirstChild.SelectSingleNode("UseClassicUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase);
            bool ribbonUIEnabled = uiControl.SettingsDocument.FirstChild.SelectSingleNode("UseRibbonUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase);
            if (classicUIEnabled)
            {
                _addDictionary.Add("$classicUICreateCall$", CreateClassicUICallCode);
                _addDictionary.Add("$classicUIRemoveCall$", RemoveClassicUICallCode);

                if (TargetProgrammLanguage == TargetProgrammingLanguage.CSharp)
                {
                    string createRemoveMethods = string.Format("{1}\t\t\t#region UserInterface{1}{1}{0}\t\t\t#endregion{1}", CreateClassicUIMethodCode + Environment.NewLine + RemoveClassicUIMethodCode, Environment.NewLine);
                    if (ribbonUIEnabled)
                        createRemoveMethods = Environment.NewLine + createRemoveMethods;
                    _addDictionary.Add("$classicUICreateRemoveMethod$", createRemoveMethods);
                }
                else
                {
                    string createRemoveMethods = string.Format("#Region \"UserInterface\"{1}{1}{0}#End Region{1}", CreateClassicUIMethodCode + Environment.NewLine + RemoveClassicUIMethodCode, Environment.NewLine);
                    if (ribbonUIEnabled)
                        createRemoveMethods = Environment.NewLine + Environment.NewLine + createRemoveMethods;
                    else
                        createRemoveMethods = Environment.NewLine + createRemoveMethods;

                    _addDictionary.Add("$classicUICreateRemoveMethod$", createRemoveMethods);
                }
            }
            else
            {
                _addDictionary.Add("$classicUICreateCall$", "");
                _addDictionary.Add("$classicUIRemoveCall$", "");
                _addDictionary.Add("$classicUICreateRemoveMethod$", "");
            }
        }

        void SetUglyRibbonUI()
        {
            IWizardControl uiControl = (ListControls[3] as IWizardControl);
            bool ribbonUIEnabled = uiControl.SettingsDocument.FirstChild.SelectSingleNode("UseRibbonUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase);
            bool classicUIEnabled = uiControl.SettingsDocument.FirstChild.SelectSingleNode("UseClassicUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase);
            if (ribbonUIEnabled)
            {
                string ribbonImplementCode = Environment.NewLine + RibbonImplementCode;
                if (!classicUIEnabled)
                    ribbonImplementCode += Environment.NewLine;
                _addDictionary.Add("$ribbonImplement$", RibbonImplement);
                _addDictionary.Add("$ribbonUIImplementMethod$", ribbonImplementCode);
                _addDictionary.Add("$ribbonFileReference$", RibbonRessourceReference);
                _addDictionary.Add("$helperCode$", RibbonHelperMethod);
            }
            else
            {
                _addDictionary.Add("$ribbonImplement$", "");
                _addDictionary.Add("$ribbonUIImplementMethod$", "");
                _addDictionary.Add("$ribbonFileReference$", "");
                _addDictionary.Add("$helperCode$", "");
            }
        }

        protected internal void RemoveRibbonRessourceFile()
        {
            IWizardControl uiControl = (ListControls[3] as IWizardControl);
            bool ribbonUIEnabled = uiControl.SettingsDocument.FirstChild.SelectSingleNode("UseRibbonUI").InnerText.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase);
            if (!ribbonUIEnabled)
            {
                string destinationAssemblyFolder = _projectFolder;
                string fileName = Path.Combine(_projectFolder, "RibbonUI.xml");
                File.Delete(fileName);
            }
        }

        #endregion

        #region Overrides

        internal override void FinishAction()
        {
            _addDictionary.Clear();
            _addDictionary.Add("$randomGuid$", Guid.NewGuid().ToString());
            _addDictionary.Add("$assemblyGuid$", Guid.NewGuid().ToString());

            string usingItems = GetDefaultUsings();

            string registerCode = "";
            string unregisterCode = "";

            foreach (XmlNode item in (ListControls[0] as IWizardControl).SettingsDocument.FirstChild.ChildNodes)
            {
                if (item.Attributes[0].Value.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
                {
                    usingItems += UsingCode.Replace("%Name%", item.Name);

                    string addinName = GetAddinName();
                    string addinDescription = GetAddinDescription();
                    string addinHiveKey = GetAddinHiveKey();

                    string addinOfficeKey = GetAddinOfficeKey(item.Name);
                    string addinLoadBehvior = GetAddinLoadBehavior();

                    string registerValue = RegisterCode;
                    registerValue = registerValue.Replace("%Name%", addinName);
                    registerValue = registerValue.Replace("%Description%", addinDescription);
                    registerValue = registerValue.Replace("%HiveKey%", addinHiveKey);
                    registerValue = registerValue.Replace("%OfficAddinKey%", addinOfficeKey);
                    registerValue = registerValue.Replace("%OfficeApp%", item.Name);
                    registerValue = registerValue.Replace("%LoadBehavior%", addinLoadBehvior);
                    registerCode += registerValue + Environment.NewLine;
                    if (_targetProgrammingLanguage == TargetProgrammingLanguage.CSharp)
                        unregisterCode += string.Format("{3}{3}{3}{3}Registry.{0}.DeleteSubKey(@\"{1}\");{2}", addinHiveKey, addinOfficeKey, Environment.NewLine, "\t");
                    else
                        unregisterCode += string.Format("{3}{3}{3}Registry.{0}.DeleteSubKey(\"{1}\"){2}", addinHiveKey, addinOfficeKey, Environment.NewLine, "\t");
                }
            }

            _addDictionary.Add("$usingItems$", usingItems);
            _addDictionary.Add("$registerCode$", registerCode);
            _addDictionary.Add("$unregisterCode$", unregisterCode);
           
            SetUglyRibbonUI();
            SetClassicUI();
            SetAssemblyReferences();
        }

        protected internal void RunStarted(Dictionary<string, string> replacementsDictionary, TargetProgrammingLanguage targetProgrammingLanguage, TargetProjectType projectType)
        {
            _targetProgrammingLanguage = targetProgrammingLanguage;
            _targetProjectType = projectType;
            _replacementsDictionary = replacementsDictionary;

            _projectFolder = replacementsDictionary["$destinationdirectory$"];
            _targetRuntime = replacementsDictionary["$targetframeworkversion$"];

            HostControl hostControl = new HostControl();
            NameControl nameControl = new NameControl();
            AddinLoadControl loadControl = new AddinLoadControl();
            AddinGuiControl guiControl = new AddinGuiControl();
            SummaryControl sumControl = new SummaryControl(this);

            ListControls.Add(hostControl);
            ListControls.Add(nameControl);
            ListControls.Add(loadControl);
            ListControls.Add(guiControl);
            ListControls.Add(sumControl);

            WizardDialog dialog = new WizardDialog(this);
            dialog.ShowDialog();

            _neededNetOfficeAssemblies = this.NeededAssemblies;

            foreach (KeyValuePair<string, string> item in this.AddDictionary)
                replacementsDictionary.Add(item.Key, item.Value);
        }

        internal override string Name
        {
            get
            {
                return "Automations Addin";
            }
        }

        #endregion
    }
}
