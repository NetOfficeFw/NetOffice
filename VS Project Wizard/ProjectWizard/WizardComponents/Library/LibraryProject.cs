using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;

namespace NetOffice.ProjectWizard
{
    public class LibraryProject : NetOfficeProject
    {
        #region Fields

        string _createClassicUIMethodCode;
        string _removeClassicUIMethodCode;
        string _createClassicUICallCode;
        string _removeClassicUICallCode;

        #endregion

        #region Properties

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
                        _removeClassicUIMethodCode = ReadString("RemoveClassicUIVisualBasic.txt");
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

        #endregion

        #region Overrides

        internal override void FinishAction()
        {
            _addDictionary.Clear();
            _addDictionary.Add("$assemblyGuid$", Guid.NewGuid().ToString());

            string usingItems = GetDefaultUsings();

            foreach (XmlNode item in (ListControls[0] as IWizardControl).SettingsDocument.FirstChild.ChildNodes)
            {
                if (item.Attributes[0].Value.Equals("TRUE", StringComparison.InvariantCultureIgnoreCase))
                    usingItems += UsingCode.Replace("%Name%", item.Name);
            }

            _addDictionary.Add("$usingItems$", usingItems);

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
            SummaryControl sumControl = new SummaryControl(this);

            ListControls.Add(hostControl);
            ListControls.Add(nameControl);
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
                return "Console Project";
            }
        }

        #endregion
    }
}
