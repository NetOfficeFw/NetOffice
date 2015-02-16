using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

using NetOffice.OutlookApi.Tools;

namespace Sample.Addin
{ 
    /// <summary>
    /// Custom pane for Outlook. The control implements the ITaskPane interface from NetOffice.Outlook.Tools
    /// </summary>
    public partial class TwitterPane : UserControl, ITaskPane
    {
        #region Ctor

        public TwitterPane()
        {
            Config = new Properties.Settings();
            InitializeComponent();
        }

        #endregion

        #region Properties

        internal static TwitterTimer Client { get; set; }
        internal static Properties.Settings Config { get; set; }

        #endregion

        #region Methods

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

        #endregion

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

        public void OnDisconnection()
        {
            Config.Save();
        }

        public void OnDockPositionChanged(NetOffice.OfficeApi.Enums.MsoCTPDockPosition position)
        {

        }

        public void OnVisibleStateChanged(bool visible)
        {

        }
    
        #endregion

        #region Trigger

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

        void Client_EnabledChanegd(bool value)
        {
            tweetGrid.Enabled = value;
        }

        void Client_Error(Exception exception)
        {
            settingsPane.SetEnabled(false);
            errorPane.ShowError(exception);
        }

        #endregion
    }
}
