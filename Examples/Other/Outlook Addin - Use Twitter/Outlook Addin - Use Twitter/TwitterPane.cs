using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Sample.Addin
{
    public partial class TwitterPane : UserControl
    {
        internal static TwitterTimer Client { get; set; }

        public TwitterPane()
        {
            InitializeComponent();
            InitializeTaskPane();
            Client = new TwitterTimer(this);
            tweetGrid.DataSource = Client;
           
          
            string result = Client.Logon();
            if(string.Empty == result)
                Client.Enabled = true;
        }

        internal void InitializeTaskPane()
        {
            tweetGrid.Visible = false;
            tweetGrid.Location = new Point(0, 0);
            tweetGrid.Size = new Size(this.Width, this.Height - buttonMain.Height);
            tweetGrid.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            buttonMain.Tag = tweetGrid;

            settingsPane.Visible = false;
            settingsPane.Location = new Point(0, 0);
            settingsPane.Size = new Size(this.Width, this.Height - buttonMain.Height);
            settingsPane.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top;
            buttonAddinSettings.Tag = settingsPane;

            button_Click(buttonMain, new EventArgs());
        }

        private void button_Click(object sender, EventArgs e)
        {
            tweetGrid.Visible = false;
            settingsPane.Visible = false;
            buttonMain.BackColor = Color.LightSteelBlue;
            buttonAddinSettings.BackColor = Color.LightSteelBlue;

            Button selectedButton = sender as Button;
            Control panel = selectedButton.Tag as Control;
            panel.Visible = true;
            selectedButton.BackColor = Color.Goldenrod;
        }
    }
}
