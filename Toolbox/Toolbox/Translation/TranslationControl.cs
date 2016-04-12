using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Currents used language changed event handler
    /// </summary>
    /// <param name="sender">translation sender</param>
    /// <param name="lcid">language id</param>
    public delegate void LanuageChangedEventHandler(object sender, int lcid);

    /// <summary>
    /// Shows all available languages
    /// </summary>
    [RessourceTable("Ressources.TranslationControlStrings.txt")]
    public partial class TranslationControl : UserControl, ILocalizationDesign
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public TranslationControl()
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = false;
        }

        #endregion

        #region Events

        /// <summary>
        /// User want close the instance
        /// </summary>
        [Category("!A")]
        public event EventHandler UserClose;

        private void RaiseUserClose()
        {
            if (null != UserClose)
                UserClose(this, EventArgs.Empty);
        }

        /// <summary>
        /// User want see the about pane
        /// </summary>
        [Category("!A")]
        public event EventHandler UserTranslationAbout;

        /// <summary>
        /// User selected another language
        /// </summary>
        [Category("!A")]
        public event LanuageChangedEventHandler LanguageChanged;

        /// <summary>
        /// User want delete a language
        /// </summary>
        [Category("!A")]
        public event LanuageChangedEventHandler LanguageDeleted;

        private void RaiseLanguageChanged(int lcid)
        {
            if (null != LanguageChanged)
                LanguageChanged(this, lcid);
        }

        private void RaiseLanguageDeleted(int lcid)
        {
            if (null != LanguageDeleted)
                LanguageDeleted(this, lcid);
        }

        #endregion

        #region Properties

        private ToolLanguage Selected
        {
            get
            {
                if (dataGridView1.SelectedCells.Count == 0)
                    return null;

                DataGridViewRow row = dataGridView1.Rows[dataGridView1.SelectedCells[0].RowIndex];
                ToolLanguage selLanguage = row.Tag as ToolLanguage;
                return selLanguage;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Show all available languages
        /// </summary>
        internal void ShowLanguages()
        {
            Forms.MainForm.Singleton.Languages.ListChanged += new ListChangedEventHandler(Languages_ListChanged);
            dataGridView1.Rows.Clear();
            foreach (var item in Forms.MainForm.Singleton.Languages)
            {
                var row = new DataGridViewRow();
                row.CreateCells(dataGridView1, new object[] { String.IsNullOrWhiteSpace(item.NameGlobal) == true ? "<Empty>" : item.NameGlobal });
                row.Tag = item;
                dataGridView1.Rows.Add(row);

            }
        }

        /// <summary>
        /// Change current used language
        /// </summary>
        /// <param name="lcid">language id</param>
        internal void SetLanguage(int lcid)
        {
            ToolLanguage language = Forms.MainForm.Singleton.Languages[lcid, false];
            if (null != language)
            {
                var component = language.Application.Components["Language Selector"];
                Translator.TranslateControls(this, component.ControlRessources);
            }
            else
            {
                Translation.Translator.TranslateControls(this, "Ressources.TranslationControlStrings.txt", lcid);
            }
        }

        private DataGridViewRow GetRow(ToolLanguage language)
        {
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                ToolLanguage lang = item.Tag as ToolLanguage;
                if (lang == language)
                    return item;
            }
            return null;
        }

        private void RaiseUserTranslationAbout()
        {
            if (null != UserTranslationAbout)
                UserTranslationAbout(this, EventArgs.Empty);
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

        public System.ComponentModel.IContainer Components
        {
            get
            {
                return components;
            }
        }

        public void Localize(string name, string text)
        {
            Translation.Translator.TranslateControl(this, name, text);
        }

        public string GetCurrentText(string name)
        {
            return Translation.Translator.TryGetControlText(this, name);
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

        private void Languages_ListChanged(object sender, ListChangedEventArgs e)
        {
            switch (e.ListChangedType)
            {
                case ListChangedType.ItemAdded:
                    {
                        ToolLanguage language = Forms.MainForm.Singleton.Languages[e.NewIndex];
                        var row = new DataGridViewRow();
                        row.CreateCells(dataGridView1, new object[] { language.NameGlobal });
                        row.Tag = language;
                        dataGridView1.Rows.Add(row);
                        break;
                    }
                case ListChangedType.ItemChanged:
                    {
                        var language = Forms.MainForm.Singleton.Languages[e.NewIndex];
                        var row = GetRow(language);
                        if (null != row)
                            row.Cells[0].Value = String.IsNullOrWhiteSpace(language.NameGlobal) == true ? "<Empty>" : language.NameGlobal;
                        break;
                    }
                case ListChangedType.ItemDeleted:
                    {
                        List<DataGridViewRow> listRows = new List<DataGridViewRow>();
                        foreach (DataGridViewRow item in dataGridView1.Rows)
                        {
                            ToolLanguage lang = item.Tag as ToolLanguage;
                            if (!Forms.MainForm.Singleton.Languages.Contains(lang))
                                listRows.Add(item);
                        }
                        foreach (var item in listRows)
                            dataGridView1.Rows.Remove(item);

                        break;
                    }
                case ListChangedType.Reset:
                    ShowLanguages();
                    break;
            }
        }

        private void toolStripClose_Click(object sender, EventArgs e)
        {
            RaiseUserClose();
        }

        private void toolStripAbout_Click(object sender, EventArgs e)
        {
            RaiseUserTranslationAbout();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (null != Selected)
            {
                bool b = ToolLanguageForm.ShowForm(this, Selected);
                if (b)
                    RaiseLanguageChanged(Selected.LCID);
            }
        }

        private void toolStripAddLanguage_Click(object sender, EventArgs e)
        {
            ToolLanguage template = Forms.SelectLanguageForm.ShowForm(this, "Select a language template");
            if (null != template)
            {
                this.Refresh();
                ToolLanguage newLanguage = new ToolLanguage(Forms.MainForm.Singleton.Languages, template);
                Forms.MainForm.Singleton.Languages.Add(newLanguage);
                if (dataGridView1.SelectedCells.Count > 0)
                    dataGridView1.SelectedCells[0].Selected = false;
                dataGridView1.Rows[dataGridView1.Rows.Count - 1].Selected = true;
                dataGridView1_DoubleClick(dataGridView1, EventArgs.Empty);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            ToolLanguage language = Selected;
            if (null != language)
                toolStripRemoveLanguage.Enabled = !(language is ToolDefaultLanguage);
            else
                toolStripRemoveLanguage.Enabled = false;
        }

        private void toolStripRemoveLanguage_Click(object sender, EventArgs e)
        {
            if (Selected is ToolDefaultLanguage)
                return;

            int lcid = Selected.LCID;
            Forms.MainForm.Singleton.Languages.Remove(Selected);
            Forms.MainForm.Singleton.Languages.ValidateFiles();
            RaiseLanguageDeleted(lcid);
        }

        #endregion
    }
}
