using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NetOffice.OutlookApi.Tools;

namespace Sample.Addin
{
    public partial class TwitterPane : UserControl, ITaskPane
    {
        internal static TwitterTimer Client { get; set; }
        internal static Properties.Settings Config { get; set; }

        public TwitterPane()
        {
            Config = new Properties.Settings();
            InitializeComponent();
        }
       
        internal void InitializeTaskPane()
        {
            tweetGrid.Visible = false;           
            buttonMain.Tag = tweetGrid;
          
            settingsPane.Visible = false;
            settingsPane.Location = new Point(0, 0);
            settingsPane.Size = new Size(this.Width, this.Height - (splitContainerButtons.Height + errorPane.Height));
            settingsPane.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            buttonAddinSettings.Tag = settingsPane;
        }

        private void button_Click(object sender, EventArgs e)
        {
            errorPane.ClearError();
            tweetGrid.Visible = false;
            settingsPane.Visible = false;
            buttonMain.BackColor = Color.LightSteelBlue;
            buttonAddinSettings.BackColor = Color.LightSteelBlue;

            Button selectedButton = sender as Button;
            Control panel = selectedButton.Tag as Control;
            panel.Visible = true;
            selectedButton.BackColor = Color.Goldenrod;
        }

        private void Client_Error(Exception exception)
        {
            settingsPane.SetEnabled(false);
            errorPane.ShowError(exception);
        }

        private void Client_EnabledChanegd(bool value)
        {
            tweetGrid.Enabled = value;
        }

        #region ITaskPane Member

        public void OnConnection(NetOffice.OutlookApi.Application application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            InitializeTaskPane();
            Client = new TwitterTimer(this, Client_Error);
            Client.EnabledChanegd += new EnabledChangedEventHanlder(Client_EnabledChanegd);
            settingsPane.Initialize(Client);
            tweetGrid.DataSource = Client;
            button_Click(buttonMain, new EventArgs());
            if (Config.Enabled)
            {
                tweetGrid.Enabled = true;
                Client.Enabled = true;
            }
            else
            {
                tweetGrid.Enabled = false;
            }
        }

        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {

        }

        public void OnVisibleStateChanged(bool visible)
        {

        }
        
        public void OnDisconnection()
        {
            Config.Save(); 
        }

        #endregion
    }
}
