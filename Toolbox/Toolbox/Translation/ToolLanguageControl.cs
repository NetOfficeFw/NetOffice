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
    /// Contains all edit language controls together
    /// </summary>
    public partial class ToolLanguageControl : UserControl
    {
        #region Fields

        private ToolLanguage _selectedLanguage;

        #endregion

        #region Ctor

        public ToolLanguageControl()
        {
            InitializeComponent();
        }

        #endregion

        #region Events

        /// <summary>
        /// The current component has been changed
        /// </summary>
        public event EventHandler SelectedNodeTextChanged;

        private void RaiseSelectedNodeTextChanged()
        {
            if (null != SelectedNodeTextChanged)
                SelectedNodeTextChanged(null, EventArgs.Empty);
        }

        /// <summary>
        /// The selected tab has been changed
        /// </summary>
        public event EventHandler SelectedTabChanged;

        private void RaiseSelectedTabChanged()
        {
            if (null != SelectedTabChanged)
                SelectedTabChanged(this, EventArgs.Empty);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Current Selected component or summary tab caption
        /// </summary>
        internal string SelectedNodeText
        {
            get
            {
                if (tabControl1.SelectedIndex == 1)
                    return languageApplicationControl1.SelectedNodeText;
                else if (tabControl1.SelectedIndex == 2)
                    return languageComponentsControl1.SelectedNodeText;
                else
                    return tabPage1.Text;
            }
        }

        /// <summary>
        /// Current selected tab index
        /// </summary>
        public int SelectedTabIndex
        {
            get
            {
                return tabControl1.SelectedIndex;
            }
        }

        /// <summary>
        /// Current edit language
        /// </summary>
        internal ToolLanguage SelectedLanguage
        {
            get
            {
                return _selectedLanguage;
            }
            set
            {
                _selectedLanguage = value;
                languageSummaryControl1.SelectedLanguage = value;
                languageApplicationControl1.SelectedLanguage = value;
                languageComponentsControl1.SelectedLanguage = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Redirect key down to the component controls
        /// </summary>
        internal void HandleKeyDown()
        {
            languageApplicationControl1.HandleKeyDown();
            languageComponentsControl1.HandleKeyDown();
        }

        /// <summary>
        /// Redirect key down to the component controls
        /// </summary>
        internal void HandleKeyUp()
        {
            languageApplicationControl1.HandleKeyUp();
            languageComponentsControl1.HandleKeyUp();
        }

        #endregion

        #region Trigger

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            RaiseSelectedNodeTextChanged();
            languageApplicationControl1.DisableComponents();
            languageComponentsControl1.DisableComponents();
            RaiseSelectedTabChanged();
        }

        private void languageApplicationControl1_SelectionChanged(object sender, EventArgs e)
        {
            RaiseSelectedNodeTextChanged();
        }

        private void languageComponentsControl1_SelectionChanged(object sender, EventArgs e)
        {
            RaiseSelectedNodeTextChanged();
        }

        #endregion
    }
}
