using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Text;
using System.IO;
using System.Net;
using System.Web;
using System.Drawing;
using Twitter = LinqToTwitter;
using LinqToTwitter;

namespace Sample.Addin
{ 
    internal class WebImage
    {
        internal WebImage(string url, Image image)
        {
            Url = url;
            Image = image;
        }

        internal string Url { get; private set; }
        internal Image Image { get; private set; }
    }

    internal class WebImageCollection 
    {
        public WebImageCollection()
        {
            Images = new List<WebImage>();
        }
        private List<WebImage> Images{get;set;}

        public Image this[string url]
        {
            get
            {
                foreach (WebImage item in Images)
	            {
                    if (item.Url.Equals(url, StringComparison.InvariantCultureIgnoreCase))
                        return item.Image;
	            }

                Image newImage = GetImageFromURL(url);
                Images.Add(new WebImage(url, newImage));
                return newImage;
            }
        }

        /// <summary>
        /// Gets the image from URL.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        private static Image GetImageFromURL(string url)
        {
            try
            {
                HttpWebRequest httpWebRequest = (HttpWebRequest)HttpWebRequest.Create(url);
                HttpWebResponse httpWebReponse = (HttpWebResponse)httpWebRequest.GetResponse();
                Stream stream = httpWebReponse.GetResponseStream();
                return Image.FromStream(stream);
            }
            catch
            {
                return null;
            }
        }
       
    }

    internal class TwitterTimer : BindingList<Twitter.Status>
    {
        #region Fields

        int _defaultInterval = 90;
        bool _enabled;
        int _interval;
        Twitter.ITwitterAuthorizer _authorizer;
        Twitter.TwitterContext _twitterContext;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="control"></param>
        public TwitterTimer(System.Windows.Forms.Control control)
        {
            Control = control;
            Interval = _defaultInterval;
        }

        #endregion

        #region Properties

        private System.Windows.Forms.Control Control{get;set;}

        public bool Enabled
        {
            get 
            {
                return _enabled;
            }
            set
            {
                _enabled = value;
                if (_enabled)
                {
                    if (null == Timer)
                    {
                        Timer = new Timer(new TimerCallback(TimerElapsed), null, 0, Interval * 1000);
                    }
                }
                else
                {
                    if (null != Timer)
                    {
                        Timer.Dispose();
                        Timer = null;
                    }
                }
            }
        }

        public int Interval
        {
            get
            {
                return _interval;
            }
            set
            {
                _interval = value;
            }
        }

        private Timer Timer { get; set; }

        #endregion
        
        #region Query Timer Trigger 
        
        List<Twitter.Status> _newList;

        private void TimerElapsed()
        {
            if (null != _newList)
            {                
                foreach (var newItem in _newList)
                {
                    bool found = false;
                    foreach (var item in this)
                    {
                        if(item.StatusID == newItem.StatusID)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (found)
                        continue;
                    this.Add(newItem);           
                }
                _newList = null;
            }
        }

        private void TimerElapsed(object val)
        {
            try
            {
                _newList = GetUserTimeLine();
                Control.Invoke(new System.Windows.Forms.MethodInvoker(TimerElapsed));
            }
            catch (Exception exception)
            {
                Enabled = false;
            }
        }
       
        #endregion

        #region Methods

        /// <summary>
        /// logon
        /// </summary>
        /// <returns></returns>
        public string Logon()
        {
            try
            {
                _authorizer = PerformAuthorization();
                _twitterContext = new Twitter.TwitterContext(_authorizer);
                return _authorizer.IsAuthorized == true ? string.Empty : "Authorization failed";
            }
            catch(Exception exception)
            {
                return exception.Message;
            }
        }


        /// <summary>
        /// create retweet
        /// </summary>
        /// <param name="tweet"></param>
        internal void CreateRetweet(Twitter.Status tweet)
        {
            _twitterContext.Retweet(tweet.StatusID);
        }

        /// <summary>
        /// tweet fav
        /// </summary>
        /// <param name="tweet"></param>
        internal void CreateFavourite(Twitter.Status tweet)
        {
           
            _twitterContext.CreateFavorite(tweet.StatusID);
        }

        /// <summary>
        /// send a new tweet
        /// </summary>
        /// <param name="text"></param>
        internal void SendTweet(string text)
        {
            _twitterContext.UpdateStatus(text);
        }

        /// <summary>
        /// logon
        /// </summary>
        /// <returns></returns>
        Twitter.ITwitterAuthorizer PerformAuthorization()
        {
            Properties.Settings settings = new Properties.Settings();

            var auth = new Twitter.SingleUserAuthorizer 
            {  
                Credentials = new Twitter.InMemoryCredentials
                {
                    OAuthToken = settings.AccessToken,
                    AccessToken = settings.AccessSecret,
                    ConsumerKey = settings.AuthenticationKey,
                    ConsumerSecret = settings.AuthenticationKey,
                },
                UseCompression = true,
            };
            auth.Authorize();
           
            return auth;
        }

        /// <summary>
        /// get new tweets
        /// </summary>
        /// <returns></returns>
        internal List<Twitter.Status> GetUserTimeLine()
        {
            var queryResponse =
                from tweet in _twitterContext.Status
                where tweet.Type == Twitter.StatusType.Home orderby tweet.CreatedAt ascending
                select tweet;
            return queryResponse.ToList();
        }

        #endregion
    }
}
