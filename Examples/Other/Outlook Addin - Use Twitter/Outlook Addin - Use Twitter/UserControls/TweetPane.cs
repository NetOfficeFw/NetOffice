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
    public delegate void TweetReplyEventHandler(Twitter.Status tweet);
    public delegate void TweetFavorizeEventHandler(Twitter.Status tweet, ref bool sucseed);
    public delegate void TweetRetweetEventHandler(Twitter.Status tweet, ref bool sucseed);

    public partial class TweetPane : UserControl
    {
        private static WebImageCollection WebImages = new WebImageCollection();

        /// <summary>
        /// windows forms designer
        /// </summary>
        public TweetPane()
        {
            InitializeComponent();
        }

        private Twitter.Status Tweet { get; set; }

        public event TweetReplyEventHandler Reply;

        private void RaiseReply()
        {
            if (null != Reply)
            {
                Reply(Tweet);
            }
        }

        public event TweetRetweetEventHandler Retweet;
        
        private void RaiseRetweet()
        {
            if (null != Retweet)
            {
                bool sucseed = false;
                Retweet(Tweet, ref sucseed);
                linkLabelRetweet.Enabled = !sucseed;
            }
        }

        public event TweetFavorizeEventHandler Favorize;

        private void RaiseFavorize()
        {
            if (null != Favorize)
            {
                bool sucseed = false;
                Favorize(Tweet, ref sucseed);
                linklabelFavorite.Enabled = !sucseed;
            }
        }

        private void linkLabelReply_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RaiseReply();
        }

        private void linkLabelRetweet_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RaiseRetweet();
        }

        private void linklabelFavorite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RaiseFavorize();
        }

        public TweetPane(Twitter.Status tweet)
        {
            InitializeComponent();
            Tweet = tweet;
            richTextBoxMessage.Text = tweet.Text;
            labelCreated.Text = tweet.CreatedAt.ToString();
            labelUserName.Text = tweet.User.Name;
            pictureBoxImage.Image =  WebImages[tweet.User.ProfileImageUrlHttps];
            linklabelFavorite.Enabled = !tweet.Favorited;
            linkLabelRetweet.Enabled = !tweet.Retweeted;
        }

        private void richTextBoxMessage_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void richTextBoxMessage_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }
    }
}
