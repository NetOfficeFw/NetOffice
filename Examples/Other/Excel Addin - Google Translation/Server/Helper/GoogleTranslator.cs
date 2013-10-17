using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    /// <summary>
    ///  taken from http://www.codeproject.com/Articles/12711/Google-Translator
    /// </summary>
    internal class GoogleTranslator : WebRessourceProvider
    {
        #region Ctor

        static GoogleTranslator()
        {
            InitializeLanguageModeMap();
        }
        
        /// <summary>
        /// Initializes a new instance of the class.
        /// </summary>
        public GoogleTranslator()
        {
            SourceLanguage = "English";
            TargetLanguage = "French";
            Referer = "http://www.google.com";
        }

        public GoogleTranslator(string sourceLanguage, string destLanguage, string sourceText)
        {
            SourceLanguage = sourceLanguage;
            TargetLanguage = destLanguage;
            SourceText = sourceText;
            Referer = "http://www.google.com";
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the source "
        /// </summary>
        /// <value>The source "</value>
        public string SourceLanguage { get; set; }

        /// <summary>
        /// Gets or sets the target "
        /// </summary>
        /// <value>The target "</value>
        public string TargetLanguage { get; set; }

        /// <summary>
        /// Gets or sets the source text.
        /// </summary>
        /// <value>The source text.</value>
        public string SourceText { get; set; }

        /// <summary>
        /// Gets the translation.
        /// </summary>
        /// <value>The translated text.</value>
        public string Translation { get; private set; }

        /// <summary>
        /// Gets the reverse translation.
        /// </summary>
        /// <value>The reverse translated text.</value>
        public string ReverseTranslation { get; private set; }
       
        #endregion        

        #region Public methods

        /// <summary>
        /// Attempts to translate the text.
        /// </summary>
        public void Translate()
        {
            // Validate source and target languages
            if (string.IsNullOrEmpty(this.SourceLanguage) ||
                string.IsNullOrEmpty(this.TargetLanguage) ||
                this.SourceLanguage.Trim().Equals(this.TargetLanguage.Trim()))
            {
                throw new Exception("An invalid source or target language was specified.");
            }

            // Delegate to base class
            base.fetchResource();
        }
     
        #endregion

        #region WebResourceProvider implementation

        /// <summary>
        /// Returns the url to be fetched.
        /// </summary>
        /// <returns>The url to be fetched.</returns>
        protected override string getFetchUrl()
        {
            return "http://translate.google.com/translate_t";
        }

        /// <summary>
        /// Retrieves the POST data (if any) to be sent to the url to be fetched.
        /// The data is returned as a string of the form "arg=val[&arg=val]...".
        /// </summary>
        /// <returns>A string containing the POST data or null if none.</returns>
        protected override string getPostData()
        {
            // Set translation mode
            string strPostData = string.Format("hl=en&ie=UTF8&oe=UTF8submit=Translate&langpair={0}|{1}",
                                                 GoogleTranslator.LanguageEnumToIdentifier(this.SourceLanguage),
                                                 GoogleTranslator.LanguageEnumToIdentifier(this.TargetLanguage));

            // Set text to be translated
            strPostData += "&text=\"" + this.SourceText + "\"";
            return strPostData;
        }          

            /// <summary>
            /// Parses the fetched content.
            /// </summary>
            protected override void parseContent()
            {
                // Initialize the scraper
                this.Translation = string.Empty;
                string strContent = this.Content;
                StringParser parser = new StringParser (strContent);

                // Scrape the translation
                string strTranslation = string.Empty;
                if (parser.skipToEndOf ("<span id=result_box")) {
                    if (parser.skipToEndOf ("onmouseout=\"this.style.backgroundColor='#fff'\">")) {
                        if (parser.extractTo("</span>", ref strTranslation)) {
                            strTranslation = StringParser.removeHtml (strTranslation);
                        }
                    }
                }

                #region Fix up the translation
                    int startClean = 0;
                    int endClean = 0;
                    int i=0;
                    while (i < strTranslation.Length) {
                        if (Char.IsLetterOrDigit (strTranslation[i])) {
                            startClean = i;
                            break;
                        }
                        i++;
                    }
                    i = strTranslation.Length - 1;
                    while (i > 0) {
                        char ch = strTranslation[i];
                        if (Char.IsLetterOrDigit (ch) ||
                            (Char.IsPunctuation (ch) && (ch != '\"'))) {
                            endClean = i;
                            break;
                        }
                        i--;
                    }
                    this.Translation = strTranslation.Substring (startClean, endClean - startClean + 1).Replace ("\"", "");
                #endregion
            }

        #endregion

        #region Private methods

        private static void InitializeLanguageModeMap()
        {
                GoogleTranslator.LanguageModeMap = new Dictionary<string, string>();
                GoogleTranslator.LanguageModeMap.Add("Afrikaans", "af");
                GoogleTranslator.LanguageModeMap.Add("Albanian", "sq");
                GoogleTranslator.LanguageModeMap.Add("Arabic", "ar");
                GoogleTranslator.LanguageModeMap.Add("Belarusian", "be");
                GoogleTranslator.LanguageModeMap.Add("Bulgarian", "bg");
                GoogleTranslator.LanguageModeMap.Add("Catalan", "ca");
                GoogleTranslator.LanguageModeMap.Add("Chinese", "zh-CN");
                GoogleTranslator.LanguageModeMap.Add("Croatian", "hr");
                GoogleTranslator.LanguageModeMap.Add("Czech", "cs");
                GoogleTranslator.LanguageModeMap.Add("Danish", "da");
                GoogleTranslator.LanguageModeMap.Add("Dutch", "nl");
                GoogleTranslator.LanguageModeMap.Add("English", "en");
                GoogleTranslator.LanguageModeMap.Add("Estonian", "et");
                GoogleTranslator.LanguageModeMap.Add("Filipino", "tl");
                GoogleTranslator.LanguageModeMap.Add("Finnish", "fi");
                GoogleTranslator.LanguageModeMap.Add("French", "fr");
                GoogleTranslator.LanguageModeMap.Add("Galician", "gl");
                GoogleTranslator.LanguageModeMap.Add("German", "de");
                GoogleTranslator.LanguageModeMap.Add("Greek", "el");
                GoogleTranslator.LanguageModeMap.Add("Haitian Creole ALPHA", "ht");
                GoogleTranslator.LanguageModeMap.Add("Hebrew", "iw");
                GoogleTranslator.LanguageModeMap.Add("Hindi", "hi");
                GoogleTranslator.LanguageModeMap.Add("Hungarian", "hu");
                GoogleTranslator.LanguageModeMap.Add("Icelandic", "is");
                GoogleTranslator.LanguageModeMap.Add("Indonesian", "id");
                GoogleTranslator.LanguageModeMap.Add("Irish", "ga");
                GoogleTranslator.LanguageModeMap.Add("Italian", "it");
                GoogleTranslator.LanguageModeMap.Add("Japanese", "ja");
                GoogleTranslator.LanguageModeMap.Add("Korean", "ko");
                GoogleTranslator.LanguageModeMap.Add("Latvian", "lv");
                GoogleTranslator.LanguageModeMap.Add("Lithuanian", "lt");
                GoogleTranslator.LanguageModeMap.Add("Macedonian", "mk");
                GoogleTranslator.LanguageModeMap.Add("Malay", "ms");
                GoogleTranslator.LanguageModeMap.Add("Maltese", "mt");
                GoogleTranslator.LanguageModeMap.Add("Norwegian", "no");
                GoogleTranslator.LanguageModeMap.Add("Persian", "fa");
                GoogleTranslator.LanguageModeMap.Add("Polish", "pl");
                GoogleTranslator.LanguageModeMap.Add("Portuguese", "pt");
                GoogleTranslator.LanguageModeMap.Add("Romanian", "ro");
                GoogleTranslator.LanguageModeMap.Add("Russian", "ru");
                GoogleTranslator.LanguageModeMap.Add("Serbian", "sr");
                GoogleTranslator.LanguageModeMap.Add("Slovak", "sk");
                GoogleTranslator.LanguageModeMap.Add("Slovenian", "sl");
                GoogleTranslator.LanguageModeMap.Add("Spanish", "es");
                GoogleTranslator.LanguageModeMap.Add("Swahili", "sw");
                GoogleTranslator.LanguageModeMap.Add("Swedish", "sv");
                GoogleTranslator.LanguageModeMap.Add("Thai", "th");
                GoogleTranslator.LanguageModeMap.Add("Turkish", "tr");
                GoogleTranslator.LanguageModeMap.Add("Ukrainian", "uk");
                GoogleTranslator.LanguageModeMap.Add("Vietnamese", "vi");
                GoogleTranslator.LanguageModeMap.Add("Welsh", "cy");
                GoogleTranslator.LanguageModeMap.Add("Yiddish", "yi");
        }

        /// <summary>
        /// Converts a language to its identifier.
        /// </summary>
        /// <param name="language">The language."</param>
        /// <returns>The identifier or <see cref="string.Empty"/> if none.</returns>
        private static string LanguageEnumToIdentifier(string language)
        {
            string mode = string.Empty;
            GoogleTranslator.LanguageModeMap.TryGetValue(language, out mode);
            return mode;              
        }

        #endregion

        #region Fields

       /// <summary>
       /// The language to translation mode map.
       /// </summary>
       internal static Dictionary<string, string> LanguageModeMap{get;set;}

       #endregion
    }
}
