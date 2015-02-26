using System;
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
        /// Translate a text
        /// </summary>
        /// <param name="sourceLanguage">source language</param>
        /// <param name="destLanguage">target language</param>
        /// <param name="text">requested text</param>
        /// <returns>A result object</returns>
        public string Translate(string sourceLanguage, string destLanguage, string text)
        {
            try
            {
                string source = null;
                string dest = null;
                if (!LanguageModeMap.TryGetValue(sourceLanguage, out source))
                    throw new ArgumentOutOfRangeException("Unkown language: " + sourceLanguage);
                if (!LanguageModeMap.TryGetValue(destLanguage, out dest))
                    throw new ArgumentOutOfRangeException("Unkown language" + destLanguage);

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

        /// <summary>
        /// The real translation job so far. (should be async in a real-life scenario)
        /// </summary>
        /// <param name="input">given text as any</param>
        /// <param name="languagePair">language set as source|dest</param>
        /// <returns>translated text</returns>
        private static string TranslateText(string input, string languagePair)
        {
            string url = String.Format("http://www.google.com/translate_t?hl=en&ie=UTF8&text={0}&langpair={1}", input, languagePair);
            WebClient webClient = new WebClient();
            webClient.Encoding = System.Text.Encoding.UTF8;
            string result = webClient.DownloadString(url);
            result = result.Substring(result.IndexOf("<span title=\"") + "<span title=\"".Length);
            result = result.Substring(result.IndexOf(">") + 1);
            result = result.Substring(0, result.IndexOf("</span>"));
            return result.Trim();
        }

        /// <summary>
        /// Initialize Language Mapping, Key value pair of Language Name, Language Code
        /// </summary>
        /// <param name="languageMap">target map to initialize</param>
        private static void InitLanguageMap(Dictionary<string, string> languageMap)
        {
            languageMap.Add("Afrikaans", "af");
            languageMap.Add("Albanian", "sq");
            languageMap.Add("Arabic", "ar");
            languageMap.Add("Armenian", "hy");
            languageMap.Add("Azerbaijani", "az");
            languageMap.Add("Basque", "eu");
            languageMap.Add("Belarusian", "be");
            languageMap.Add("Bengali", "bn");
            languageMap.Add("Bulgarian", "bg");
            languageMap.Add("Catalan", "ca");
            languageMap.Add("Chinese", "zh-CN");
            languageMap.Add("Croatian", "hr");
            languageMap.Add("Czech", "cs");
            languageMap.Add("Danish", "da");
            languageMap.Add("Dutch", "nl");
            languageMap.Add("English", "en");
            languageMap.Add("Esperanto", "eo");
            languageMap.Add("Estonian", "et");
            languageMap.Add("Filipino", "tl");
            languageMap.Add("Finnish", "fi");
            languageMap.Add("French", "fr");
            languageMap.Add("Galician", "gl");
            languageMap.Add("German", "de");
            languageMap.Add("Georgian", "ka");
            languageMap.Add("Greek", "el");
            languageMap.Add("Haitian Creole", "ht");
            languageMap.Add("Hebrew", "iw");
            languageMap.Add("Hindi", "hi");
            languageMap.Add("Hungarian", "hu");
            languageMap.Add("Icelandic", "is");
            languageMap.Add("Indonesian", "id");
            languageMap.Add("Irish", "ga");
            languageMap.Add("Italian", "it");
            languageMap.Add("Japanese", "ja");
            languageMap.Add("Korean", "ko");
            languageMap.Add("Lao", "lo");
            languageMap.Add("Latin", "la");
            languageMap.Add("Latvian", "lv");
            languageMap.Add("Lithuanian", "lt");
            languageMap.Add("Macedonian", "mk");
            languageMap.Add("Malay", "ms");
            languageMap.Add("Maltese", "mt");
            languageMap.Add("Norwegian", "no");
            languageMap.Add("Persian", "fa");
            languageMap.Add("Polish", "pl");
            languageMap.Add("Portuguese", "pt");
            languageMap.Add("Romanian", "ro");
            languageMap.Add("Russian", "ru");
            languageMap.Add("Serbian", "sr");
            languageMap.Add("Slovak", "sk");
            languageMap.Add("Slovenian", "sl");
            languageMap.Add("Spanish", "es");
            languageMap.Add("Swahili", "sw");
            languageMap.Add("Swedish", "sv");
            languageMap.Add("Tamil", "ta");
            languageMap.Add("Telugu", "te");
            languageMap.Add("Thai", "th");
            languageMap.Add("Turkish", "tr");
            languageMap.Add("Ukrainian", "uk");
            languageMap.Add("Urdu", "ur");
            languageMap.Add("Vietnamese", "vi");
            languageMap.Add("Welsh", "cy");
            languageMap.Add("Yiddish", "yi");
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
