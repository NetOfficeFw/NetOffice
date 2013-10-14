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
    public partial class OpenFilePanelDialog : Form
    {
        public OpenFilePanelDialog()
        {
            InitializeComponent();
        }

        public OpenFilePanelDialog(OpenFilePanelDialogLanguage language = OpenFilePanelDialogLanguage.English)
        {
            InitializeComponent();
        }

        public OpenFilePanel Panel
        {
            get
            {
                return InnerOpenFilePanel;
            }
        }

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

        public string SelectedFile
        {
            get
            {
                return InnerOpenFilePanel.SelectedFile;
            }
        }

        public string[] SelectedFiles
        {
            get
            {
                return InnerOpenFilePanel.SelectedFiles;
            }
        }

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
    }
}
