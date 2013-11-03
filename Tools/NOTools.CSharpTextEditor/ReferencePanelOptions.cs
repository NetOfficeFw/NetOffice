using System;
using System.Drawing;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.CSharpTextEditor
{
    /// <summary>
    ///All possible settings for the reference panel
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class ReferencePanelOptions
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">Parent editor control</param>
        internal ReferencePanelOptions(CodeEditorControl parent)
        {
            Parent = parent;
            InitLocalization();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Parent editor control
        /// </summary>
        private CodeEditorControl Parent { get; set; }

        /// <summary>
        /// Allow the user to see the error panel
        /// </summary>
        [DisplayName("AllowPanel"), Category("CodeEditor"), Description("Allow the user to see the reference panel")]
        public bool AllowPanel
        {
            get
            {
                return !Parent.splitContainer3.Panel2Collapsed;
            }
            set
            {
                Parent.splitContainer3.Panel2Collapsed = !value;
            }
        }

        /// <summary>
        /// Allow the user to add/remove references
        /// </summary>
        [DisplayName("AllowAddRemove"), Category("CodeEditor"), Description("Allow the user to add/remove references")]
        public bool AllowAddRemoveReferences { get; set; }

        /// <summary>
        /// Allow the user to add Non-GAC references from the local file system
        /// </summary>
        [DisplayName("AllowFileReferences"), Category("CodeEditor"), Description("Allow the user to add Non-GAC references from the local file system")]
        public bool AllowAddFileReferences { get; set; }

        /// <summary>
        /// GAC Tab Label
        /// </summary>
        [DisplayName("GACTitle"), Category("Localization"), Description("GAC Tab Label")]
        public string GACTitle { get; set; }

        /// <summary>
        /// FileSystem Tab Label
        /// </summary>
        [DisplayName("FileSystemTitle"), Category("Localization"), Description("FileSystem Tab Label")]
        public string FileSystemTitle { get; set; }

        /// <summary>
        /// Add Menu Item Label
        /// </summary>
        [DisplayName("AddTitle"), Category("Localization"), Description("Add Menu Item Label")]
        public string AddTitle { get; set; }

        /// <summary>
        /// Remove Menu Item Label
        /// </summary>
        [DisplayName("RemoveTitle"), Category("Localization"), Description("Remove Menu Item Label")]
        public string RemoveTitle { get; set; }

        /// <summary>
        /// Ok Button Label
        /// </summary>
        [DisplayName("OkButtonTitle"), Category("Localization"), Description("Ok Button Label")]
        public string OkButtonTitle { get; set; }

        /// <summary>
        /// Cancel Button Label
        /// </summary>
        [DisplayName("CancelButtonTitle"), Category("Localization"), Description("Cancel Button Label")]
        public string CancelButtonTitle { get; set; }

        /// <summary>
        /// DialogTitle Text Label
        /// </summary>
        [DisplayName("DialogTitle"), Category("Localization"), Description("DialogTitle Text Label")]
        public string DialogTitle { get; set; }

        /// <summary>
        /// Get or set the error panel is open
        /// </summary>
        [Category("CodeEditor"), Description("Get or set the reference panel is open")]
        public bool PanelOpen
        {
            get
            {
                return Parent.referencePanel1.PanelOpen;
            }
            set
            {
                Parent.referencePanel1.PanelOpen = value;
            }
        }

        /// <summary>
        /// Header message for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Header message for the reference panel")]
        public string Header
        {
            get
            {
                return Parent.referencePanel1.labelHeader.Text;
            }
            set
            {
                Parent.referencePanel1.labelHeader.Text = value;
            }
        }

        /// <summary>
        /// Back color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Back color for the reference panel")]
        public Color BackColor
        {
            get
            {
                return Parent.referencePanel1.BackColor;
            }
            set
            {
                Parent.referencePanel1.BackColor = value;
            }
        }

        /// <summary>
        /// Fore color for the error panel
        /// </summary>
        [Category("CodeEditor"), Description("Fore color for the reference panel")]
        public Color ForeColor
        {
            get
            {
                return Parent.referencePanel1.ForeColor;
            }
            set
            {
                Parent.referencePanel1.ForeColor = value;
            }
        }
       
        #endregion

        #region Methods

        private void InitLocalization()
        {
            GACTitle = "GAC";
            FileSystemTitle = "FileSystem";
            AddTitle = "Add";
            RemoveTitle = "Remove";
            OkButtonTitle = "Ok";
            CancelButtonTitle = "Cancel";
            DialogTitle = "Choose Reference";
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String instance that represents the class instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return "ReferencePanelOptions";
        }
        
        #endregion
    }
}
