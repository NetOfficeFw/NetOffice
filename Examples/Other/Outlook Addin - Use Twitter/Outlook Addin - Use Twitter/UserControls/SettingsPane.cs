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
        public SettingsPane()
        {
            InitializeComponent();
            Initialize();
        }

        public void Initialize()
        {
            InitializeState = true;
            Config = new Properties.Settings();       
           
            LoadConfiguration();
            InitializeState = false;
            
        }

        internal Properties.Settings Config { get; set; }
        private bool InitializeState { get; set; }         

        internal void LoadConfiguration()
        {
            textBoxAuthKey.Text = Config.AuthenticationKey;
            textBoxAuthSecret.Text = Config.AuthenticationSecret;
            textBoxAccessToken.Text = Config.AccessToken;
            textBoxAccessSecret.Text = Config.AccessSecret;
            numericRefreshInterval.Value = Config.RefreshInterval;
            textBoxTweetAlert.Text = Config.Alerts;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Config.Save();
        }

        private void textBoxAuthKey_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.AuthenticationKey = textBoxAuthKey.Text;
        }

        private void textBoxAuthSecret_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.AuthenticationSecret = textBoxAuthSecret.Text;
        }

        private void textBoxAccessToken_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.AccessToken = textBoxAccessToken.Text;
        }

        private void textBoxAccessSecret_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.AccessSecret = textBoxAccessSecret.Text;
        }

        private void numericRefreshInterval_ValueChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.RefreshInterval =numericRefreshInterval.Value;
        }

        private void textBoxTweetAlert_TextChanged(object sender, EventArgs e)
        {
            if (InitializeState)
                return;
            Config.Alerts = textBoxTweetAlert.Text;
        }

        private void buttonTestConnection_Click(object sender, EventArgs e)
        {

        }
    }
}
