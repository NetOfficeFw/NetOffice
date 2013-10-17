using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    /// <summary>
    ///  taken from http://www.codeproject.com/Articles/12711/Google-Translator
    /// </summary>
    class StringParser
    {
        private string m_strContent = "";
        private string m_strContentLC = "";
        private int m_nIndex = 0;
        public string Content
        {
            get
            {
                return this.m_strContent;
            }
            set
            {
                this.m_strContent = value;
                this.m_strContentLC = this.m_strContent.ToLower();
                this.resetPosition();
            }
        }
        public int Position
        {
            get
            {
                return this.m_nIndex;
            }
        }
        public StringParser()
        {
        }
        public StringParser(string strContent)
        {
            this.Content = strContent;
        }
        public static void getLinks(string strString, string strRootUrl, ref ArrayList documents, ref ArrayList images)
        {
            strString = StringParser.removeComments(strString);
            strString = StringParser.removeScripts(strString);
            StringParser stringParser = new StringParser(strString);
            stringParser.replaceEvery("'", "\"");
            string text = "";
            if (strRootUrl != null)
            {
                text = strRootUrl.Trim();
            }
            if (text.Length > 0 && !text.EndsWith("/"))
            {
                text += "/";
            }
            string text2 = "";
            stringParser.resetPosition();
            while (stringParser.skipToEndOfNoCase("href=\""))
            {
                if (stringParser.extractTo("\"", ref text2))
                {
                    text2 = text2.Trim();
                    if (text2.Length > 0 && text2.IndexOf("mailto:") == -1)
                    {
                        if (!text2.StartsWith("http://") && !text2.StartsWith("ftp://"))
                        {
                            try
                            {
                                UriBuilder uriBuilder = new UriBuilder(text);
                                uriBuilder.Path =text2;
                                text2 = uriBuilder.Uri.ToString();
                            }
                            catch (Exception)
                            {
                                text2 = "http://" + text2;
                            }
                        }
                        if (!documents.Contains(text2))
                        {
                            documents.Add(text2);
                        }
                    }
                }
            }
            stringParser.resetPosition();
            while (stringParser.skipToEndOfNoCase("src=\""))
            {
                if (stringParser.extractTo("\"", ref text2))
                {
                    text2 = text2.Trim();
                    if (text2.Length > 0)
                    {
                        if (!text2.StartsWith("http://") && !text2.StartsWith("ftp://"))
                        {
                            try
                            {
                                UriBuilder uriBuilder2 = new UriBuilder(text);
                                uriBuilder2.Path = text2;
                                text2 = uriBuilder2.Uri.ToString();
                            }
                            catch (Exception)
                            {
                                text2 = "http://" + text2;
                            }
                        }
                        if (!images.Contains(text2))
                        {
                            images.Add(text2);
                        }
                    }
                }
            }
        }
        public static string removeComments(string strString)
        {
            string text = "";
            string text2 = "";
            StringParser stringParser = new StringParser(strString);
            while (stringParser.extractTo("<!--", ref text2))
            {
                text += text2;
                if (!stringParser.skipToEndOf("-->"))
                {
                    return strString;
                }
            }
            stringParser.extractToEnd(ref text2);
            return text + text2;
        }
        public static string removeEnclosingAnchorTag(string strString)
        {
            string text = strString.ToLower();
            int num = text.IndexOf("<a");
            if (num != -1)
            {
                num++;
                num = text.IndexOf(">", num);
                if (num != -1)
                {
                    num++;
                    int num2 = text.LastIndexOf("</a>");
                    if (num2 != -1)
                    {
                        return strString.Substring(num, num2 - num);
                    }
                }
            }
            return strString;
        }
        public static string removeEnclosingQuotes(string strString)
        {
            int num = strString.IndexOf("\"");
            if (num != -1)
            {
                int num2 = strString.LastIndexOf("\"");
                if (num2 > num)
                {
                    return strString.Substring(num, num2 - num - 1);
                }
            }
            return strString;
        }
        public static string removeHtml(string strString)
        {
            Hashtable hashtable = new Hashtable();
            hashtable.Add("&nbsp;", " ");
            hashtable.Add("&amp;", "&");
            hashtable.Add("&aring;", "");
            hashtable.Add("&auml;", "");
            hashtable.Add("&eacute;", "");
            hashtable.Add("&iacute;", "");
            hashtable.Add("&igrave;", "");
            hashtable.Add("&ograve;", "");
            hashtable.Add("&ouml;", "");
            hashtable.Add("&quot;", "\"");
            hashtable.Add("&szlig;", "");
            StringParser stringParser = new StringParser(strString);
            IEnumerator enumerator = hashtable.Keys.GetEnumerator();
            try
            {
                while (enumerator.MoveNext())
                {
                    string text = (string)enumerator.Current;
                    string strReplacement = hashtable[text] as string;
                    if (strString.IndexOf(text) != -1)
                    {
                        stringParser.replaceEveryExact(text, strReplacement);
                    }
                }
            }
            finally
            {
                IDisposable disposable = enumerator as IDisposable;
                if (disposable != null)
                {
                    disposable.Dispose();
                }
            }
            stringParser.replaceEveryExact("&#0", "&#");
            stringParser.replaceEveryExact("&#39;", "'");
            stringParser.replaceEveryExact("</", " <~/");
            stringParser.replaceEveryExact("<~/", "</");
            hashtable.Clear();
            hashtable.Add("<br>", " ");
            hashtable.Add("<p>", " ");
            enumerator = hashtable.Keys.GetEnumerator();
            try
            {
                while (enumerator.MoveNext())
                {
                    string text2 = (string)enumerator.Current;
                    string strReplacement2 = hashtable[text2] as string;
                    if (strString.IndexOf(text2) != -1)
                    {
                        stringParser.replaceEvery(text2, strReplacement2);
                    }
                }
            }
            finally
            {
                IDisposable disposable = enumerator as IDisposable;
                if (disposable != null)
                {
                    disposable.Dispose();
                }
            }
            strString = stringParser.Content;
            string text3 = "";
            int num = 0;
            int num2;
            while ((num2 = strString.IndexOf("<", num)) != -1)
            {
                string text4 = strString.Substring(num, num2 - num);
                text3 += text4;
                num = num2 + 1;
                int num3 = strString.IndexOf(">", num);
                if (num3 == -1)
                {
                    break;
                }
                num = num3 + 1;
            }
            if (num < strString.Length)
            {
                text3 += strString.Substring(num, strString.Length - num);
            }
            strString = text3;
            stringParser.Content = strString;
            stringParser.replaceEveryExact("  ", " ");
            strString = stringParser.Content.Trim();
            return strString;
        }
        public static string removeScripts(string strString)
        {
            string text = "";
            string text2 = "";
            StringParser stringParser = new StringParser(strString);
            while (stringParser.extractToNoCase("<script", ref text2))
            {
                text += text2;
                if (!stringParser.skipToEndOfNoCase("</script>"))
                {
                    stringParser.Content = text;
                    return strString;
                }
            }
            stringParser.extractToEnd(ref text2);
            return text + text2;
        }
        public bool at(string strString)
        {
            return this.m_strContent.IndexOf(strString, this.Position) == this.Position;
        }
        public bool atNoCase(string strString)
        {
            strString = strString.ToLower();
            return this.m_strContentLC.IndexOf(strString, this.Position) == this.Position;
        }
        public bool extractTo(string strString, ref string strExtract)
        {
            int num = this.m_strContent.IndexOf(strString, this.Position);
            if (num != -1)
            {
                strExtract = this.m_strContent.Substring(this.m_nIndex, num - this.m_nIndex);
                this.m_nIndex = num + strString.Length;
                return true;
            }
            return false;
        }
        public bool extractToNoCase(string strString, ref string strExtract)
        {
            strString = strString.ToLower();
            int num = this.m_strContentLC.IndexOf(strString, this.Position);
            if (num != -1)
            {
                strExtract = this.m_strContent.Substring(this.m_nIndex, num - this.m_nIndex);
                this.m_nIndex = num + strString.Length;
                return true;
            }
            return false;
        }
        public bool extractUntil(string strString, ref string strExtract)
        {
            int num = this.m_strContent.IndexOf(strString, this.Position);
            if (num != -1)
            {
                strExtract = this.m_strContent.Substring(this.m_nIndex, num - this.m_nIndex);
                this.m_nIndex = num;
                return true;
            }
            return false;
        }
        public bool extractUntilNoCase(string strString, ref string strExtract)
        {
            strString = strString.ToLower();
            int num = this.m_strContentLC.IndexOf(strString, this.Position);
            if (num != -1)
            {
                strExtract = this.m_strContent.Substring(this.m_nIndex, num - this.m_nIndex);
                this.m_nIndex = num;
                return true;
            }
            return false;
        }
        public void extractToEnd(ref string strExtract)
        {
            strExtract = "";
            if (this.Position < this.m_strContent.Length)
            {
                int num = this.m_strContent.Length - this.Position;
                strExtract = this.m_strContent.Substring(this.Position, num);
            }
        }
        public int replaceEvery(string strOccurrence, string strReplacement)
        {
            int num = 0;
            strOccurrence = strOccurrence.ToLower();
            for (int num2 = this.m_strContentLC.IndexOf(strOccurrence); num2 != -1; num2 = this.m_strContentLC.IndexOf(strOccurrence))
            {
                string text = this.m_strContent.Substring(0, num2) + strReplacement;
                int num3 = num2 + strOccurrence.Length;
                if (num3 < this.m_strContent.Length)
                {
                    string text2 = this.m_strContent.Substring(num3, this.m_strContent.Length - num3);
                    text += text2;
                }
                this.m_strContent = text;
                this.m_strContentLC = this.m_strContent.ToLower();
                num++;
            }
            return num;
        }
        public int replaceEveryExact(string strOccurrence, string strReplacement)
        {
            int num = 0;
            while (this.m_strContent.IndexOf(strOccurrence) != -1)
            {
                this.m_strContent = this.m_strContent.Replace(strOccurrence, strReplacement);
                num++;
            }
            this.m_strContentLC = this.m_strContent.ToLower();
            return num;
        }
        public void resetPosition()
        {
            this.m_nIndex = 0;
        }
        public bool skipToStartOf(string strString)
        {
            return this.seekTo(strString, false, false);
        }
        public bool skipToStartOfNoCase(string strText)
        {
            return this.seekTo(strText, true, false);
        }
        public bool skipToEndOf(string strString)
        {
            return this.seekTo(strString, false, true);
        }
        public bool skipToEndOfNoCase(string strText)
        {
            return this.seekTo(strText, true, true);
        }
        private bool seekTo(string strString, bool bNoCase, bool bPositionAfter)
        {
            if (this.Position >= this.m_strContent.Length)
            {
                return false;
            }
            int num;
            if (bNoCase)
            {
                strString = strString.ToLower();
                num = this.m_strContentLC.IndexOf(strString, this.Position);
            }
            else
            {
                num = this.m_strContent.IndexOf(strString, this.Position);
            }
            if (num == -1)
            {
                return false;
            }
            this.m_nIndex = num;
            if (bPositionAfter)
            {
                this.m_nIndex += strString.Length;
            }
            return true;
        }
    }
}
