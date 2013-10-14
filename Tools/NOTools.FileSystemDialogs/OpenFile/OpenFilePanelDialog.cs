using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// OpenFilePanel Wrapper Dialog
    /// </summary>
    public partial class OpenFilePanelDialog : Form
    {
        #region Ctor
        
        public OpenFilePanelDialog()
        {
            InitializeComponent();
        }

        public OpenFilePanelDialog(OpenFilePanelDialogLanguage language = OpenFilePanelDialogLanguage.English)
        {
            InitializeComponent();
        }

        #endregion
        
        #region Properties

        /// <summary>
        /// Inner wrapped panel
        /// </summary>
        public OpenFilePanel Panel
        {
            get
            {
                return InnerOpenFilePanel;
            }
        }

        /// <summary>
        /// Filter(for example "All(*.*)|*.*")
        /// </summary>
        public string FileFilter
        {
            get
            {
                return InnerOpenFilePanel.Misc.FileFilter;
            }
            set
            {
                InnerOpenFilePanel.Misc.FileFilter = value;
            }
        }

        /// <summary>
        /// User selected file
        /// </summary>
        public string SelectedFile
        {
            get
            {
                return InnerOpenFilePanel.SelectedFile;
            }
        }

        /// <summary>
        /// User selected files (if multiselect allowed)
        /// </summary>
        public string[] SelectedFiles
        {
            get
            {
                return InnerOpenFilePanel.SelectedFiles;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Simple static Show method for easy use
        /// </summary>
        /// <param name="parent">parent window</param>
        /// <param name="fileFilter">file filter (for example "All(*.*)|*.*")</param>
        /// <param name="language">additional language setting</param>
        /// <returns>Dialog result Ok Or Cancel with selected file</returns>
        public static StaticOpenFilePanelDialogResult Show(IWin32Window parent, string fileFilter,OpenFilePanelDialogLanguage language = OpenFilePanelDialogLanguage.English)
        {
            if (null == parent)
                throw new ArgumentNullException();
            OpenFilePanelDialog dialog = new OpenFilePanelDialog(language);           
            dialog.InnerOpenFilePanel.Misc.FileFilter = fileFilter;
            DialogResult dr =  dialog.ShowDialog(parent);
            return new StaticOpenFilePanelDialogResult(dr, dialog.InnerOpenFilePanel.SelectedFiles);
        }

        private void SetLanguage(OpenFilePanelDialogLanguage language)
        {
            if (language == OpenFilePanelDialogLanguage.German)
            {
                this.Text = "Datei öffnen";
                ButtonSelect.Text = "Öffnen";
                ButtonCancel.Text = "Abbrechen";
                InnerOpenFilePanel.Localization.Set1031Default(this, new EventArgs());
            }
        }

        #endregion

        #region Trigger

        private void ButtonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void ButtonSelect_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void InnerOpenFilePanel_SelectionChanged(object sender, SelectionChangedEventArgs args)
        {
            ButtonSelect.Enabled = args.Files.Length > 0;
        }

        private void InnerOpenFilePanel_FileDoubleClick(object sender, FileDoubleClickEventArgs args)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}
