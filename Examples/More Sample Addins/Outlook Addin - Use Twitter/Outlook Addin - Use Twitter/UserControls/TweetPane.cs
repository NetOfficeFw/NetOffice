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
   
    public partial class TweetPane : UserControl
    {
        #region Properties

        /// <summary>
        /// Image Cache
        /// </summary>
        private static WebImageCollection WebImages = new WebImageCollection();
        
        /// <summary>
        /// Handled Tweet
        /// </summary>
        private Twitter.Status Tweet { get; set; }
      
        #endregion

        #region Ctor

        /// <summary>
        /// Creates instance from class - windows forms designer ctor
        /// </summary>
        public TweetPane()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Creates instance from class 
        /// </summary>
        /// <param name="tweet">handled tweet</param>
        public TweetPane(Twitter.Status tweet)
        {
            InitializeComponent();
            Tweet = tweet;
            richTextBoxMessage.Text = tweet.Text;
            labelCreated.Text = tweet.CreatedAt.ToString();
            labelUserName.Text = tweet.User.Name;
            pictureBoxImage.Image = WebImages[tweet.User.ProfileImageUrlHttps];
            linklabelFavorite.Enabled = !tweet.Favorited;
            linkLabelRetweet.Enabled = !tweet.Retweeted;
        }
        #endregion
         
        #region Events

        /// <summary>
        /// Signal to the pane the user want reply to the message
        /// </summary>
        public event TweetReplyEventHandler Reply;

        private void RaiseReply()
        {
            if (null != Reply)
            {
                Reply(Tweet);
            }
        }

        #endregion

        #region UI Trigger

        private void linkLabelReply_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            RaiseReply();
        }

        private void linkLabelRetweet_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabelRetweet.Enabled = TwitterPane.Client.CreateRetweet(Tweet);
        }

        private void linklabelFavorite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linklabelFavorite.Enabled = TwitterPane.Client.CreateFavourite(Tweet);
        }

        private void richTextBoxMessage_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true;
        }

        private void richTextBoxMessage_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch
            {
                // no worries
            }

        }

        #endregion
    }
}
