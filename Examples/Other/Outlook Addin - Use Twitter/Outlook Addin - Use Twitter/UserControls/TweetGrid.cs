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
        public TweetGrid()
        {
            InitializeComponent();
            Panes = new List<TweetPane>();
        }

        BindingList<Twitter.Status> _dataSource;
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

        private List<TweetPane> Panes { get; set; }

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
                    newPane.Size = new Size(panelTweetPanels.Width, 115);
                    newPane.Visible = true;
                    foreach (var item in Panes)
                        item.Location = new Point(item.Location.X, item.Location.Y + 115);
                    Panes.Add(newPane);
                    panelTweetPanels.Controls.Add(newPane);
                    vScrollBar.Maximum = Panes.Count *115;
                    newPane.Reply += new TweetReplyEventHandler(newPane_Reply);
                    newPane.Favorize += new TweetFavorizeEventHandler(newPane_Favorize);
                    newPane.Retweet += new TweetRetweetEventHandler(newPane_Retweet);
                    break;
                }
                default:
                    break;
            }            
        }

        void newPane_Retweet(Twitter.Status tweet, ref bool sucseed)
        {
            try
            {
                TwitterPane.Client.CreateRetweet(tweet);
                sucseed = true;
            }
            catch (Exception)
            {
                sucseed = false;
            }
            
        }

        void newPane_Favorize(Twitter.Status tweet, ref bool sucseed)
        {
            try
            {
                TwitterPane.Client.CreateFavourite(tweet);
                sucseed = true;
            }
            catch (Exception)
            {
                sucseed = false;                
            }
        }

        void newPane_Reply(Twitter.Status tweet)
        {
            textBoxTweetContent.Text = "@" + tweet.User.Identifier.ScreenName;
            textBoxTweetContent.Focus();
        }

        private void vScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            int shift = e.NewValue - e.OldValue;
            foreach (var item in Panes)
            {
                item.Top -= shift;
            }
        }

        private void buttonSendTweet_Click(object sender, EventArgs e)
        {
            try
            {
                panelError.Visible = false;
                TwitterPane.Client.SendTweet(textBoxTweetContent.Text);
                textBoxTweetContent.Clear();
            }
            catch (Exception exception)
            {
                panelError.Visible = true;
                labelErrorMessage.Text = exception.Message;
            }
        }

        private void textBoxTweetContent_TextChanged(object sender, EventArgs e)
        {
            buttonSendTweet.Enabled = !String.IsNullOrWhiteSpace(textBoxTweetContent.Text);
             
        }
    }
}
