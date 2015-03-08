using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Show about info to the user
    /// </summary>
    public partial class AboutDialog : ToolsDialog
    {
        #region Fields

        private const string _emptyValue = "<Empty>";
        private string _headerCaption;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public AboutDialog()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="headerCaption">header text to display</param>
        /// <param name="assemblyTitle">assembly title</param>
        /// <param name="assemblyVersion">assembly version</param>
        /// <param name="copyrightHint">copyright info</param>
        /// <param name="companyTextUrl">company text url</param>
        /// <param name="companyUrl">company url</param>
        /// <param name="licenceText">licence info</param>
        public AboutDialog(string headerCaption, string assemblyTitle, string assemblyVersion, string copyrightHint, string companyTextUrl, string companyUrl, string licenceText)
        {
            InitializeComponent();
            _headerCaption = String.IsNullOrEmpty(headerCaption) ? _emptyValue : headerCaption;
            labelAbout.Text = String.IsNullOrEmpty(headerCaption) ? _emptyValue : headerCaption;
            labelAssemblyTitleVersion.Text = String.Format("{0} {1}", assemblyTitle, String.IsNullOrWhiteSpace(assemblyVersion) ? _emptyValue : assemblyVersion);
            labelCopyright.Text = String.IsNullOrWhiteSpace(copyrightHint) ? _emptyValue : copyrightHint;
            if (!String.IsNullOrEmpty(companyUrl))
            {
                linkLabelCompany.Text = String.IsNullOrEmpty(companyTextUrl) ? companyUrl : companyTextUrl;
                linkLabelCompany.Tag = companyUrl;
            }
            else
            {
                linkLabelCompany.Visible = false;
            }
            richTextBoxLicence.Text = licenceText;
        }

        #endregion

        #region Methods

        private void TryProceedLink(string link)
        {
            try
            {
                if(!String.IsNullOrEmpty(link))
                    Process.Start(link);
            }
            catch
            {
                ;
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// <see cref="ToolsDialog.DoLocalization"/>
        /// </summary>
        /// <param name="localization">localized values</param>
        protected internal override void DoLocalization(DialogLocalization localization)
        {
            Text = String.Format("{0} {1}", localization["Title", Text], _headerCaption);
            labelLicenceHeader.Text = localization["LicenceHeader", labelLicenceHeader.Text];
            buttonClose.Text = localization["buttonClose", buttonClose.Text];
        }

        /// <summary>
        /// <see cref="ToolsDialog.DoLayout"/>
        /// </summary>
        /// <param name="layout">layout settings</param>
        protected internal override void DoLayout(DialogLayoutSettings layout)
        {
            panelLicenceHeader.BackColor = layout.BackAlternateColor;
            base.DoLayout(layout);
        }

        #endregion

        #region Trigger

        private void buttonClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        private void linkLabelCompany_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                TryProceedLink(linkLabelCompany.Tag as string);
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        private void richTextBoxLicence_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                TryProceedLink(e.LinkText);
            }
            catch (Exception exception)
            {
                ShowSingleException(exception);
            }
        }

        #endregion
    }
}
