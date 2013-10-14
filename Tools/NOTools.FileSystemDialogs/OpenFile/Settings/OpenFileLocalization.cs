using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class OpenFileLocalization : INotifyPropertyChanged
    {
        #region Ctor

        internal OpenFileLocalization(PropertyChangedEventHandler eventHandler = null)
        {
            PropertyBag = new PropertyBagCollection<string>("<Empty>", RaisePropertyChanged);
            Set1033Default(null, new EventArgs());
            if (null != eventHandler)
                this.PropertyChanged += eventHandler;
        }

        #endregion

        #region Properties

        [Description("Desktop Category"), Category("Localization")]
        public string Desktop
        {
            get { return PropertyBag["Desktop"]; }
            set { PropertyBag["Desktop"] = value; }
        }

        [Description("MyMachine Category"), Category("Localization")]
        public string MyMachine
        {
            get { return PropertyBag["MyMachine"]; }
            set { PropertyBag["MyMachine"] = value; }
        }

        [Description("MyDocuments Category"), Category("Localization")]
        public string MyDocuments
        {
            get { return PropertyBag["MyDocuments"]; }
            set { PropertyBag["MyDocuments"] = value; }
        }

        [Description("SpecialFolders Category"), Category("Localization")]
        public string SpecialFolders
        {
            get { return PropertyBag["SpecialFolders"]; }
            set { PropertyBag["SpecialFolders"] = value; }
        }

        [Description("Template Folders Category"), Category("Localization")]
        public string TemplateFolders
        {
            get { return PropertyBag["TemplateFolders"]; }
            set { PropertyBag["TemplateFolders"] = value; }
        }

        [Description("Selected Filename Header"), Category("Localization")]
        public string LabelFileName
        {
            get { return PropertyBag["LabelFileName"]; }
            set { PropertyBag["LabelFileName"] = value; }
        }

        [Description("Filte Filter Extension Header"), Category("Localization")]
        public string LabelFileFilter
        {
            get { return PropertyBag["LabelFileFilter"]; }
            set { PropertyBag["LabelFileFilter"] = value; }
        }

        [Description("Tooltip text for large icon view button"), Category("Localization")]
        public string LabelLargeIconView
        {
            get { return PropertyBag["LabelLargeIconView"]; }
            set { PropertyBag["LabelLargeIconView"] = value; }
        }

        [Description("Tooltip text for small icon view button"), Category("Localization")]
        public string LabelSmallIconView
        {
            get { return PropertyBag["LabelSmallIconView"]; }
            set { PropertyBag["LabelSmallIconView"] = value; }
        }

        [Description("Tooltip text for detail view button"), Category("Localization")]
        public string LabelDetailsView
        {
            get { return PropertyBag["LabelDetailsView"]; }
            set { PropertyBag["LabelDetailsView"] = value; }
        }

        [Description("Name template for new directory name"), Category("Localization")]
        public string NewDirectoryName
        {
            get { return PropertyBag["NewDirectoryName"]; }
            set { PropertyBag["NewDirectoryName"] = value; }
        }

        [Description("Tooltip text for create directory button"), Category("Localization")]
        public string LabelCreateDirectory
        {
            get { return PropertyBag["LabelCreateDirectory"]; }
            set { PropertyBag["LabelCreateDirectory"] = value; }
        }

        [Description("Tooltip text for delete directory button"), Category("Localization")]
        public string LabelDeleteDirectory
        {
            get { return PropertyBag["LabelDeleteDirectory"]; }
            set { PropertyBag["LabelDeleteDirectory"] = value; }
        }

        [Description("Tooltip text for delete file button"), Category("Localization")]
        public string LabelDeleteFile
        {
            get { return PropertyBag["LabelDeleteFile"]; }
            set { PropertyBag["LabelDeleteFile"] = value; }
        }

        [Description("Tooltip text for go upward button"), Category("Localization")]
        public string LabelGoUpward
        {
            get { return PropertyBag["LabelGoUpward"]; }
            set { PropertyBag["LabelGoUpward"] = value; }
        }

        [Description("Tooltip text for go undo button"), Category("Localization")]
        public string LabelGoUndo
        {
            get { return PropertyBag["LabelGoUndo"]; }
            set { PropertyBag["LabelGoUndo"] = value; }
        }

        [Description("Tooltip text for go redo button"), Category("Localization")]
        public string LabelGoRedo
        {
            get { return PropertyBag["LabelGoRedo"]; }
            set { PropertyBag["LabelGoRedo"] = value; }
        }

        [Description("Confirm header(messagebox) before directory delete"), Category("Localization")]
        public string AskBeforeDeleteDirectoryHeader
        {
            get { return PropertyBag["AskBeforeDeleteDirectoryHeader"]; }
            set { PropertyBag["AskBeforeDeleteDirectoryHeader"] = value; }
        }

        [Description("Confirm message before directory delete"), Category("Localization")]
        public string AskBeforeDeleteDirectoryMessage
        {
            get { return PropertyBag["AskBeforeDeleteDirectoryMessage"]; }
            set { PropertyBag["AskBeforeDeleteDirectoryMessage"] = value; }
        }

        [Description("Confirm header(messagebox) before file delete"), Category("Localization")]
        public string AskBeforeDeleteFileHeader
        {
            get { return PropertyBag["AskBeforeDeleteFileHeader"]; }
            set { PropertyBag["AskBeforeDeleteFileHeader"] = value; }
        }

        [Description("Confirm message before file delete"), Category("Localization")]
        public string AskBeforeDeleteFileMessage
        {
            get { return PropertyBag["AskBeforeDeleteFileMessage"]; }
            set { PropertyBag["AskBeforeDeleteFileMessage"] = value; }
        }

        private PropertyBagCollection<string> PropertyBag { get; set; }

        private string Default1031
        {
            get 
            {
                if (null == _default1031)
                    _default1031 = ReadRessourceTextFile("NOTools.FileSystemDialogs.OpenFile.Ressource.1031.txt");
                return _default1031;
            }
        }
        private string _default1031;

        private string Default1033
        {
            get
            {
                if (null == _default1033)
                    _default1033 = ReadRessourceTextFile("NOTools.FileSystemDialogs.OpenFile.Ressource.1033.txt");
                return _default1033;
            }
        }
        private string _default1033;

        #endregion
        
        #region Methods

        internal void Set1031Default(object sender, System.EventArgs e)
        {
            PropertyBag.Clear();
            string[] lines = Default1031.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                int position = line.IndexOf("=");
                string name = line.Substring(0, position);
                string value = line.Substring(position+1);
                PropertyBag.Add(name, value);
            }
            RaisePropertyChanged("");
        }

        internal void Set1033Default(object sender, System.EventArgs e)
        {
            PropertyBag.Clear();
            string[] lines = Default1033.Split(new string[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                int position = line.IndexOf("=");
                string name = line.Substring(0, position);
                string value = line.Substring(position + 1);
                PropertyBag.Add(name, value);
            }
            RaisePropertyChanged("");
        }

        private string ReadRessourceTextFile(string fileName)
        {
            Assembly assembly = this.GetType().Assembly;
            System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(fileName);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource File."));

            string text = textStreamReader.ReadToEnd();
            ressourceStream.Close();
            textStreamReader.Close();
            return text;
        }

        #endregion

        #region INotifyPropertyChanged

        [Browsable(false)]
        public event PropertyChangedEventHandler PropertyChanged;

        internal void RaisePropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Overrides

        public override string ToString()
        {
            return "Localization";
        }

        #endregion
    }
}
