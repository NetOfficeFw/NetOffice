using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Threading;

namespace Sample.Server
{
    /// <summary>
    ///  taken from http://www.codeproject.com/Articles/12711/Google-Translator
    /// </summary>
    abstract class WebRessourceProvider
    {
        private string m_strAgent;
		private string m_strReferer;
		private string m_strError;
		private string m_strContent;
		private HttpStatusCode m_httpStatusCode;
		private int m_nPause;
		private int m_nTimeout;
		private DateTime m_tmFetchTime;

		public string Agent
		{
			get
			{
				return this.m_strAgent;
			}
			set
			{
				this.m_strAgent = ((value == null) ? "" : value);
			}
		}
		public string Referer
		{
			get
			{
				return this.m_strReferer;
			}
			set
			{
				this.m_strReferer = ((value == null) ? "" : value);
			}
		}
		public int Pause
		{
			get
			{
				return this.m_nPause;
			}
			set
			{
				this.m_nPause = value;
			}
		}
		public int Timeout
		{
			get
			{
				return this.m_nTimeout;
			}
			set
			{
				this.m_nTimeout = value;
			}
		}
		public string Content
		{
			get
			{
				return this.m_strContent;
			}
		}
		public DateTime FetchTime
		{
			get
			{
				return this.m_tmFetchTime;
			}
		}
		public string ErrorMsg
		{
			get
			{
				return this.m_strError;
			}
		}

        public WebRessourceProvider()
		{
			this.reset();
		}

		public void reset()
		{
			this.m_strAgent = "Mozilla/4.0 (compatible; MSIE 5.5; Windows NT 5.0)";
			this.m_strReferer = "";
			this.m_strError = "";
			this.m_strContent = "";
			this.m_httpStatusCode = HttpStatusCode.OK;
			this.m_nPause = 0;
			this.m_nTimeout = 0;
			this.m_tmFetchTime = DateTime.MinValue;
		}

		public void fetchResource()
		{
			if (!this.init())
			{
				return;
			}
			bool flag;
			do
			{
				string fetchUrl = this.getFetchUrl();
				this.getContent(fetchUrl);
				flag = (this.m_httpStatusCode == HttpStatusCode.OK);
				if (flag)
				{
					this.parseContent();
				}
			}
			while (flag && this.continueFetching());
		}
		protected virtual bool init()
		{
			return true;
		}
		protected abstract string getFetchUrl();
		protected virtual string getPostData()
		{
			return null;
		}
		protected virtual void parseContent()
		{
		}
		protected virtual bool continueFetching()
		{
			return false;
		}
		private void getContent(string url)
		{
			if (this.m_nPause > 0)
			{
				int num = 0;
				do
				{
					if (num == 0 && this.m_tmFetchTime != DateTime.MinValue)
					{
						num = (int)(this.m_tmFetchTime - DateTime.Now).TotalMilliseconds;
					}
					int num2 = 100;
					if (num < this.m_nPause)
					{
						Thread.Sleep(num2);
						num += num2;
					}
				}
				while (num < this.m_nPause);
			}
			string text = url;
			if (!text.StartsWith("http://"))
			{
				text = "http://" + text;
			}
			HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(text);
			httpWebRequest.AllowAutoRedirect = true;
			httpWebRequest.UserAgent = this.m_strAgent;
			httpWebRequest.Referer = this.m_strReferer;
			if (this.m_nTimeout != 0)
			{
				httpWebRequest.Timeout = this.m_nTimeout;
			}
			string postData = this.getPostData();
			if (postData != null)
			{
				ASCIIEncoding aSCIIEncoding = new ASCIIEncoding();
				byte[] bytes = aSCIIEncoding.GetBytes(postData);
				httpWebRequest.Method = "POST";
				httpWebRequest.ContentType = "application/x-www-form-urlencoded";
				httpWebRequest.ContentLength = (long)bytes.Length;
				Stream requestStream = httpWebRequest.GetRequestStream();
				requestStream.Write(bytes, 0, bytes.Length);
				requestStream.Close();
			}
			this.m_strError = "";
			this.m_strContent = "";
			HttpWebResponse httpWebResponse = null;
			try
			{
				this.m_tmFetchTime = DateTime.Now;
				httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
			}
			catch (Exception ex)
			{
				if (ex is WebException)
				{
					WebException ex2 = ex as WebException;
					this.m_strError = ex2.Message;
				}
				return;
			}
			finally
			{
				if (httpWebResponse != null)
				{
					this.m_httpStatusCode = httpWebResponse.StatusCode;
				}
			}
			try
			{
				Stream responseStream = httpWebResponse.GetResponseStream();
				StreamReader streamReader = new StreamReader(responseStream);
				this.m_strContent = streamReader.ReadToEnd();
			}
			catch (Exception)
			{
			}
		}
    }
}
