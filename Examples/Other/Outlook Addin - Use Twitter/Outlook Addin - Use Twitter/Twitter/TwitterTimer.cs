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
    public delegate void ErrorEventHandler(Exception exception);

    public delegate void EnabledChangedEventHanlder(bool value);

    internal class TwitterTimer : BindingList<Twitter.Status>
    {
        #region Fields

        int _defaultInterval = 90;
        Twitter.ITwitterAuthorizer _authorizer;
        Twitter.TwitterContext _twitterContext;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="control"></param>
        public TwitterTimer(System.Windows.Forms.Control control, ErrorEventHandler errorHandler)
        {
            Control = control;
            IntervalSeconds = _defaultInterval;
            if (null != errorHandler)
                OperationError += errorHandler;

            // read config
            InitializeState = true;

            ConsumerKey = TwitterPane.Config.AuthenticationKey;
            ConsumerSecret = TwitterPane.Config.AuthenticationSecret;
            AccessToken = TwitterPane.Config.AccessToken;
            AccessSecret = TwitterPane.Config.AccessSecret;
            IntervalSeconds = TwitterPane.Config.RefreshInterval;
            Enabled = TwitterPane.Config.Enabled;

            InitializeState = false;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Enable or disable the Timer
        /// </summary>
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
                        if(Logon())
                            Timer = new Timer(new TimerCallback(TimerElapsed), null, 0, (int)IntervalSeconds * 1000);
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

                RaiseEnabledChanegd(value);

                if (!InitializeState)
                    TwitterPane.Config.Enabled = value;
            }
        }
        bool _enabled;

        /// <summary>
        /// Get or set the timer interval. Miminum is 90
        /// </summary>
        public decimal IntervalSeconds
        {
            get
            {
                return _intervalSeconds;
            }
            set
            {
                if (value < 90)
                    throw new ArgumentOutOfRangeException();
                _intervalSeconds = value;
                if (!InitializeState)
                    TwitterPane.Config.RefreshInterval = value;
            }
        }
        decimal _intervalSeconds;

        /// <summary>
        /// oAuth Key
        /// </summary>
        public string ConsumerKey
        {
            get
            {
                return _consumerKey;
            }
            set
            {
                _consumerKey = value;
                if (!InitializeState)
                    TwitterPane.Config.AuthenticationKey = value;
            }
        }
        string _consumerKey;

        /// <summary>
        /// oAuth Secret
        /// </summary>
        public string ConsumerSecret
        {
            get
            {
                return _consumerSecret;
            }
            set
            {
                _consumerSecret = value;
                if (!InitializeState)
                    TwitterPane.Config.AuthenticationSecret = value;
            }
        }
        string _consumerSecret;

        /// <summary>
        /// Access token
        /// </summary>
        public string AccessToken
        {
            get
            {
                return _accessToken;
            }
            set
            {
                _accessToken = value;
                if (!InitializeState)
                    TwitterPane.Config.AccessToken = value;
            }
        }
        string _accessToken;

        /// <summary>
        /// Access secret
        /// </summary>
        public string AccessSecret
        {
            get
            {
                return _accessSecret;
            }
            set
            {
                _accessSecret = value;
                if (!InitializeState)
                    TwitterPane.Config.AccessSecret = value;
            }
        }
        string _accessSecret;

        internal bool InitializeState { get; set; }


        /// <summary>
        /// Parent Control for Invoke Check
        /// </summary>
        private System.Windows.Forms.Control Control{get;set;}

        private Timer Timer { get; set; }

        #endregion
        
        #region Query Timer Trigger 
        
        List<Twitter.Status> _newList;

        private void TimerElapsed()
        {
            try
            {
                if (null != _newList)
                {
                    foreach (var newItem in _newList)
                    {
                        bool found = false;
                        foreach (var item in this)
                        {
                            if (item.StatusID == newItem.StatusID)
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
            catch (Exception exception)
            {
                RaiseOperationError(exception);
                Enabled = false;
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
                Control.Invoke(new ErrorEventHandler(RaiseOperationError), new object[] { exception });
                Enabled = false;
            }
        }
       
        #endregion

        #region Events 
        
        /// <summary>
        /// Signals the client an operation error
        /// </summary>
        public event ErrorEventHandler OperationError;

        private void RaiseOperationError(Exception exception)
        {
            if (null != OperationError)
                OperationError(exception);
        }

        public event EnabledChangedEventHanlder EnabledChanegd;

        private void RaiseEnabledChanegd(bool value)
        {
            if (null != EnabledChanegd)
                EnabledChanegd(value);
        }

        #endregion

        #region Methods

        /// <summary>
        /// logon
        /// </summary>
        /// <returns></returns>
        public bool Logon()
        {
            try
            {
                _authorizer = PerformAuthorization();
                _twitterContext = new Twitter.TwitterContext(_authorizer);

                try
                {
                    var accounts = from acct in _twitterContext.Account
                                   where acct.Type == AccountType.VerifyCredentials
                                   select acct;

                    Account account = accounts.SingleOrDefault();
                    User user = account.User;
                    Status tweet = user.Status ?? new Status();
                }
                catch (Exception exception)
                {                    
                    throw new Exception("Authentication failed.", exception);
                }

                return true;
            }
            catch(Exception exception)
            {
                RaiseOperationError(exception);
                return false;
            }
        }
        
   

        /// <summary>
        /// logon with specific credentials
        /// </summary>
        /// <returns></returns>
        public bool Logon(string accessToken, string accessSecret, string authenticationKey, string autenthicationSecret)
        {
            try
            {
                _authorizer = PerformAuthorization(accessToken, accessSecret, authenticationKey, autenthicationSecret);
                _twitterContext = new Twitter.TwitterContext(_authorizer);
               
                var accounts = from acct in _twitterContext.Account
                               where acct.Type == AccountType.VerifyCredentials
                               select acct;

                Account account = accounts.SingleOrDefault();
                User user = account.User;
                Status tweet = user.Status ?? new Status();

                
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// create retweet
        /// </summary>
        /// <param name="tweet"></param>
        internal bool CreateRetweet(Twitter.Status tweet)
        {
            try
            {
                _twitterContext.Retweet(tweet.StatusID);
                return true;
            }
            catch(Exception exception)
            {
                RaiseOperationError(exception);
                return false;
            }
        }

        /// <summary>
        /// Tweet fav
        /// </summary>
        /// <param name="tweet"></param>
        internal bool CreateFavourite(Twitter.Status tweet)
        {
            try
            {
                _twitterContext.CreateFavorite(tweet.StatusID);
                return true;
            }
            catch (Exception exception)
            {
                RaiseOperationError(exception);
                return false;
            }
        }

        /// <summary>
        /// Send a new tweet
        /// </summary>
        /// <param name="text"></param>
        internal bool SendTweet(string text)
        {
            try
            {
                _twitterContext.UpdateStatus(text);
                return true;
            }
            catch (Exception exception)
            {
                RaiseOperationError(exception);
                return false;
            }
        }

        /// <summary>
        /// Logon
        /// </summary>
        /// <returns></returns>
        private Twitter.ITwitterAuthorizer PerformAuthorization()
        {
            Properties.Settings settings = new Properties.Settings();

            var auth = new Twitter.SingleUserAuthorizer 
            {  
                Credentials = new Twitter.InMemoryCredentials
                {
                    OAuthToken = AccessToken,
                    AccessToken = AccessSecret,
                    ConsumerKey = ConsumerKey,
                    ConsumerSecret = ConsumerSecret,
                },
                UseCompression = true,
            };
            auth.Authorize();
           
            return auth;
        }

        /// <summary>
        /// Logon
        /// </summary>
        /// <returns></returns>
        private Twitter.ITwitterAuthorizer PerformAuthorization(string accessToken, string accessSecret, string authenticationKey, string autenthicationSecret)
        {
            Properties.Settings settings = new Properties.Settings();

            var auth = new Twitter.SingleUserAuthorizer
            {
                Credentials = new Twitter.InMemoryCredentials
                {
                    OAuthToken = accessToken,
                    AccessToken = accessSecret,
                    ConsumerKey = authenticationKey,
                    ConsumerSecret = autenthicationSecret,
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
        private List<Twitter.Status> GetUserTimeLine()
        {
            try
            {
                var queryResponse =
                from tweet in _twitterContext.Status
                where tweet.Type == Twitter.StatusType.Home
                orderby tweet.CreatedAt ascending
                select tweet;
                return queryResponse.ToList();
            }
            catch (Exception exception)
            {
                throw new Exception("Unable to request the Timeline.", exception);
            }
        }

        #endregion
    }
}
