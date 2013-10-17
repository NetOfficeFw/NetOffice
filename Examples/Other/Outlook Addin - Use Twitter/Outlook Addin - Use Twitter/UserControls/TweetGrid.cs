using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Twitter = LinqToTwitter;

namespace Sample.Addin
{
    public partial class TweetGrid : UserControl
    {
        BindingList<Twitter.Status> _dataSource;

        public TweetGrid()
        {
            InitializeComponent();
            panelDisabled.Dock = DockStyle.Fill;
            splitContainer1.Dock = DockStyle.Fill;
            Panes = new List<TweetPane>();
        }

        public new bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {
                _enabled = value;
                if (Enabled)
                {
                    panelDisabled.Visible = false;
                }
                else
                {
                    panelDisabled.Visible = true;
                }
            }
        }
        bool _enabled;

        private List<TweetPane> Panes { get; set; }

        public BindingList<Twitter.Status> DataSource
        {
            get
            {
                return _dataSource;
            }
            set
            {
                if(null != _dataSource)
                    _dataSource.ListChanged -= new ListChangedEventHandler(DataSource_ListChanged);
                _dataSource = value;
                if(null != _dataSource)
                    _dataSource.ListChanged += new ListChangedEventHandler(DataSource_ListChanged);
            }
        }

        void DataSource_ListChanged(object sender, ListChangedEventArgs e)
        {
            switch (e.ListChangedType)
            {
                case ListChangedType.ItemAdded:
                {
                    vScrollBar.Value = 0;
                    Twitter.Status newTweet = DataSource[e.NewIndex];

                    TweetPane newPane = new TweetPane(newTweet);
                    newPane.Location = new Point(0, 0);
                    newPane.Size = new Size(panelTweetPanels.Width, 140);
                    newPane.Visible = true;
                    foreach (var item in Panes)
                        item.Location = new Point(item.Location.X, item.Location.Y + 140);
                    Panes.Add(newPane);
                    panelTweetPanels.Controls.Add(newPane);
                    vScrollBar.Maximum = Panes.Count * 140;
                    newPane.Reply += new TweetReplyEventHandler(TweetPane_Reply);
                    break;
                }
                default:
                    break;
            }            
        }

        void TweetPane_Reply(Twitter.Status tweet)
        {
            textBoxTweetContent.Text = "@" + tweet.User.Identifier.ScreenName;
            textBoxTweetContent.Focus();
        }

        private void ScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            int shift = e.NewValue - e.OldValue;
            foreach (var item in Panes)
                item.Top -= (shift);
        }

        private void buttonSendTweet_Click(object sender, EventArgs e)
        {
            if(TwitterPane.Client.SendTweet(textBoxTweetContent.Text))
                textBoxTweetContent.Clear();
        }

        private void textBoxTweetContent_TextChanged(object sender, EventArgs e)
        {
            buttonSendTweet.Enabled = !String.IsNullOrWhiteSpace(textBoxTweetContent.Text);
        }

        private void buttonErrorDetails_Click(object sender, EventArgs e)
        {
            pongPanel.StartGame();
        }

        private void linkLabelNetOffice_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://twitter.com/net_office");
            }
            catch
            {
                ;
            }
        }
    }
}
