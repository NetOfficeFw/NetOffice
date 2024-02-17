using System;
using System.Net;
using System.Collections.Generic;
using Word = NetOffice.WordApi;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
using NetOffice.OfficeApi.Enums;

namespace Word05AddinCS4
{
    /// <summary>
    /// Deobfuscate all bit.ly URL's in a document
    /// 
    /// Please note: it is a sample implementation,
    /// real code want handle more url shorteners and is aware of cascading shorteners (and possible endless cascade).
    /// Moreover professional addins handle remote web communication outside(exe/service) of MS-Office
    /// to deal friendly with desktop firewalls.
    /// </summary>
    public class CustomDocumentInspector : ToolsDocumentInspector
    {
        private static string[] _bitly = new string[] { "http://bit.ly", "https://bit.ly" };
        private static string[] _linkEnds = new string[] { " ", "^t", "\r" };

        public CustomDocumentInspector()
        {
            InspectResult = new Dictionary<int, string>();
            Cache = new Dictionary<string, string>();
        }

        /// <summary>
        /// What we found in Inspect
        /// </summary>
        private Dictionary<int, string> InspectResult { get; set; }

        /// <summary>
        /// Short/Long Url Cache to avoid resolve the same short link twice
        /// </summary>
        private Dictionary<string, string> Cache { get; set; }

        public override string Name
        {
            get
            {
                return "Deobfuscate bit.ly Short Links(CS4)";
            }
        }

        public override string Description
        {
            get
            {
                return "Performs HTTP calls to resolve bit.ly url's.";
            }
        }
      
        public override void Inspect(Word.Document doc, out MsoDocInspectorStatus status, out string result, out string action)
        {
            InspectResult.Clear();
            Cache.Clear();
            Word.Range range = doc.Content;
            Word.Find find = range.Find;
            find.Forward = true;
            find.Text = "http*";
            find.MatchWildcards = true;
            if (find.Execute())
            {
                int start = range.Start;
                while (start > 0)
                {
                    string text = String.Empty;
                    Word.Range character = range.Characters[1];
                    while (null != character)
                    {
                        string characterText = character.Text;
                        bool isEndLink = false;
                        foreach (string item in _linkEnds)
                        {
                            if (characterText == item)
                            {
                                isEndLink = true;
                                break;
                            }
                        }
                        if (!isEndLink)
                        {
                            text += character.Text;
                            character = character.Next();
                        }
                        else
                            break;
                    }
                   
                    foreach (string item in _bitly)
                    {
                        if (text.StartsWith(item))
                        {
                            InspectResult.Add(start, text);
                            break;
                        }
                    }
                    if (!find.Execute())
                        break;
                    start = range.Start;
                }

            }
            
            if (InspectResult.Count  > 0)
            {
                status = MsoDocInspectorStatus.msoDocInspectorStatusIssueFound;
                result = String.Format("{0} link(s) found.", InspectResult.Count);
                action = "Deobfuscate Links.";
            }
            else
            {
                status = MsoDocInspectorStatus.msoDocInspectorStatusDocOk;
                result = "No links found.";
                action = "No links to change.";
            }
        }

        public override void Fix(Word.Document doc, int hwnd, out MsoDocInspectorStatus status, out string result)
        {
            Word.Range range = doc.Content;
            Word.Find find = range.Find;
            
            int replacedLinks = 0;
            foreach (KeyValuePair<int, string> item in InspectResult)
            {
                string uri = TryGetBitlyRedirectUrl(item.Value);
                if(!String.IsNullOrWhiteSpace(uri))
                {
                    if (find.Execute(item.Value, null, null, null, null, null, null, null, null, uri))
                        replacedLinks++;
                }
            }

            if (replacedLinks == InspectResult.Count)
            {
                status = MsoDocInspectorStatus.msoDocInspectorStatusDocOk;
                result = "All links have been replaced.";
            }
            else
            {
                status = MsoDocInspectorStatus.msoDocInspectorStatusError;
                result = "Unable to replace one or more link(s).";
            }
        }

        private string TryGetBitlyRedirectUrl(string uri)
        {
            if (Cache.ContainsKey(uri))
                return Cache[uri];

            HttpWebRequest request = WebRequest.Create(uri) as HttpWebRequest;
            request.Timeout = 5000;
            request.Method = "HEAD";
            request.AllowAutoRedirect = false;
            HttpWebResponse response = null;
            try
            {
                response = request.GetResponse() as HttpWebResponse;
                if (null != response)
                { 
                    string result = response.GetResponseHeader("Location");
                    if (!String.IsNullOrWhiteSpace(result) && result.EndsWith("/"))
                        result = result.Substring(0, result.Length - 1);
                    Cache.Add(uri, result);
                    return result;
                }
                else
                    return null;
            }
            catch (WebException)
            {
                // 404 - invalid bit.ly link or timeout because network issues
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}