using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace Sample.Server
{
    /// <summary>
    /// Encapsulate a translation operation
    /// </summary>
    /// <param name="result">operation result incl. state</param>
    public delegate void TranslationEventHandler(TranslateOperationResult result);

    /// <summary>
    /// Offers Google Translation Functionality
    /// </summary>
    public class WebTranslationService : MarshalByRefObject
    {
        #region Fields

        private object _lock = new object();
        private static DataEventRepeators _repeators;
        private static string[] _availableTranslations;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public WebTranslationService()
        {
            if (null == LanguageModeMap)
            {
                LanguageModeMap = new Dictionary<string, string>();
                InitLanguageMap(LanguageModeMap);
            }
            Cache = new TranslationCache();
        }

        #endregion

        #region Properties

        /// <summary>
        /// The language to translation mode map.
        /// </summary>
        internal static Dictionary<string, string> LanguageModeMap { get; set; }

        /// <summary>
        /// Local Cache
        /// </summary>
        private TranslationCache Cache { get; set; }

        /// <summary>
        /// Available Language Names
        /// </summary>
        public string[] AvailableTranslations 
        {
            get
            {
                if (null == _availableTranslations)
                {
                    List<string> list = new List<string>();
                    foreach (var item in LanguageModeMap)
                        list.Add(item.Key);
                    _availableTranslations = list.ToArray();
                }

                return _availableTranslations;
            }
        }
        
        /// <summary>
        /// Current Event Repeators
        /// </summary>
        public static DataEventRepeators Repeators
        {
            get
            {
                if (_repeators == null)
                {
                    _repeators = new DataEventRepeators();
                }
                return _repeators;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add a repeator to the instance
        /// </summary>
        /// <param name="repeater">new event repeator</param>
        public void AddEventRepeater(DataEventRepeator repeater)
        {
            Repeators.Add(repeater);
        }

        /// <summary>
        /// Translate a text (should be async in a real-life scenario)
        /// </summary>
        /// <param name="sourceLanguage">source language</param>
        /// <param name="destLanguage">target language</param>
        /// <param name="text">requested text</param>
        /// <returns>A result object</returns>
        public string Translate(string sourceLanguage, string destLanguage, string text)
        {
            if (String.IsNullOrWhiteSpace(text))
                return text;

            lock (_lock)
            {
                try
                {
                    string source = null;
                    string dest = null;
                    if (!LanguageModeMap.TryGetValue(sourceLanguage, out source))
                        throw new ArgumentOutOfRangeException("Unkown language: " + sourceLanguage);
                    if (!LanguageModeMap.TryGetValue(destLanguage, out dest))
                        throw new ArgumentOutOfRangeException("Unkown language" + destLanguage);

                    text = text.Trim();

                    LocalTranslationCacheItem cacheItem = Cache.TryGetValue(source, dest, text);
                    if (null != cacheItem)
                    {
                        TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Sucseed, text, cacheItem.TranslationText, null, true);
                        RaiseOnTranslation(result);
                        return cacheItem.TranslationText;
                    }
                    else
                    {
                        string translatedText = TranslateText(text, String.Format("{0}|{1}", source, dest));
                        TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Sucseed, text, translatedText, null);
                        Cache.Add(source, dest, text, translatedText);
                        RaiseOnTranslation(result);
                        return translatedText;
                    }
                }
                catch (Exception exception)
                {
                    TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Error, text, null, exception);
                    RaiseOnTranslation(result);
                    throw exception;
                }
            }
        }

        /// <summary>
        /// The real translation job so far
        /// </summary>
        /// <param name="input">given text as any</param>
        /// <param name="languagePair">language set as source|dest</param>
        /// <returns>translated text</returns>
        private static string TranslateText(string input, string languagePair)
        {
            if (String.IsNullOrWhiteSpace(input))
                return String.Empty;
            string url = String.Format("http://www.google.com/translate_t?hl=en&ie=UTF8&text={0}&langpair={1}", input, languagePair);
            using (WebClient webClient = new WebClient())
            {
                webClient.Encoding = System.Text.Encoding.UTF8;
                webClient.Headers.Add(HttpRequestHeader.UserAgent, "Mozilla/5.0");
                webClient.Headers.Add(HttpRequestHeader.AcceptCharset, "UTF-8");
                string result = webClient.DownloadString(url);
                return ProceedTranslationResult(input, result);
            }
        }

        /// <summary>
        /// Extract translation result from web response
        /// </summary>
        /// <param name="input">input as fallback if its failed to extract</param>
        /// <param name="result">google translation response</param>
        /// <returns>translation result or input</returns>
        private static string ProceedTranslationResult(string input, string response)
        {
            string result = response;

            string firstTarget = "TRANSLATED_TEXT='";
            string secondTarget = "'";

            int index = result.IndexOf(firstTarget, 0, StringComparison.InvariantCultureIgnoreCase);
            if (index < 0)
                return input;

            result = result.Substring(index + firstTarget.Length);
            index = result.IndexOf(secondTarget, 0, StringComparison.InvariantCultureIgnoreCase);
            if (index < 0)
                return input;

            result = result.Substring(0, index).Replace("\\r\\x3cbr\\x3e", Environment.NewLine);
            return result;
        }

        /// <summary>
        /// Initialize Language Mapping, Key value pair of Language Name, Language Code
        /// </summary>
        /// <param name="languageMap">target map to initialize</param>
        private static void InitLanguageMap(Dictionary<string, string> languageMap)
        {
            Assembly assembly = typeof(WebTranslationService).Assembly;
            Stream stream = assembly.GetManifestResourceStream(typeof(WebTranslationService).Namespace + ".Languages.txt");
            StreamReader reader = new StreamReader(stream);
            string[] languages = reader.ReadToEnd().Split(new string[]{Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries);

            foreach (string item in languages)
            {
                string[] keyPair = item.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                languageMap.Add(keyPair[0], keyPair[1]);
            }

            reader.Dispose();
            stream.Dispose();
        }

        #endregion

        #region Events

        private void RaiseOnTranslation(TranslateOperationResult result)
        {
            foreach (var item in Repeators)
            {
                try
                {
                    item.OnTranslation(result);
                }
                catch (Exception exception)
                {
                    System.Diagnostics.Trace.WriteLine("Exception occured :" + exception.Message);
                }            
            }
        }

        #endregion
    }  
}
