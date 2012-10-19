using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sample.Addin
{
    public partial class SettingsPane : UserControl
    {
        #region Properties

        private bool InitializeState { get; set; }

        #endregion

        #region Ctor

        public SettingsPane()
        {
            InitializeComponent();
        }

        #endregion

        #region UI Trigger

        private void checkBoxEnabled_CheckedChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            TwitterPane.Client.Enabled = checkBoxEnabled.Checked;
        }

        private void textBoxAuthKey_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            checkBoxEnabled.Checked = false;
            TwitterPane.Client.ConsumerKey = textBoxAuthKey.Text;
        }

        private void textBoxAuthSecret_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            checkBoxEnabled.Checked = false;
            TwitterPane.Client.ConsumerSecret = textBoxAuthSecret.Text;
        }

        private void textBoxAccessToken_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            checkBoxEnabled.Checked = false;
            TwitterPane.Client.AccessToken = textBoxAccessToken.Text;
        }

        private void textBoxAccessSecret_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            checkBoxEnabled.Checked = false;
            TwitterPane.Client.AccessSecret = textBoxAccessSecret.Text;
        }

        private void numericRefreshInterval_ValueChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            checkBoxEnabled.Checked = false;
            TwitterPane.Client.IntervalSeconds = numericRefreshInterval.Value;
        }

        private void buttonTestConnection_Click(object sender, EventArgs e)
        {
            try
            {
                labelTestConnection.Visible = false;

                TwitterTimer testTimer = new TwitterTimer(this, null);
                if (testTimer.Logon(textBoxAccessToken.Text, textBoxAccessSecret.Text, textBoxAuthKey.Text, textBoxAuthSecret.Text))
                {
                    labelTestConnection.ForeColor = Color.Green;
                    labelTestConnection.Text = "Authentication passed.";
                    labelTestConnection.Visible = true;
                }
                else
                {
                    labelTestConnection.ForeColor = Color.Red;
                    labelTestConnection.Text = "Authentication failed.";
                    labelTestConnection.Visible = true;
                }
            }
            catch
            {
                labelTestConnection.ForeColor = Color.Red;
                labelTestConnection.Text = "Authentication failed.";
                labelTestConnection.Visible = true;
            }
        }

        private void linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                LinkLabel label = sender as LinkLabel;
                System.Diagnostics.Process.Start(label.Text);
            }
            catch
            {
                ;
            }
        }

        #endregion

        #region Methods

        internal void Initialize(TwitterTimer timer)
        {
            InitializeState = true;
            LoadConfiguration(timer);
            InitializeState = false;
        }

        internal void LoadConfiguration(TwitterTimer timer)
        {
            textBoxAuthKey.Text = timer.ConsumerKey;
            textBoxAuthSecret.Text = timer.ConsumerSecret;
            textBoxAccessToken.Text = timer.AccessToken;
            textBoxAccessSecret.Text = timer.AccessSecret;
            numericRefreshInterval.Value = timer.IntervalSeconds;
            checkBoxEnabled.Checked = timer.Enabled;
        }

        internal void SetEnabled(bool value)
        {
            checkBoxEnabled.Checked = value;
        }

        #endregion
    }
}
