using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    /// <summary>
    /// A very very simple cache list
    /// </summary>
    public class TranslationCache : List<LocalTranslationCacheItem>
    {
        /// <summary>
        /// Add a new item
        /// </summary>
        /// <param name="sourceLanguage"></param>
        /// <param name="destLanguage"></param>
        /// <param name="sourceText"></param>
        /// <param name="translationText"></param>
        public void Add(string sourceLanguage, string destLanguage, string sourceText, string translationText)
        {
            this.Add(new LocalTranslationCacheItem(sourceLanguage, destLanguage, sourceText, translationText));
        }

        /// <summary>
        /// Get an existing item if present
        /// </summary>
        /// <param name="sourceLanguage">Source language</param>
        /// <param name="destLanguage">Target language</param>
        /// <param name="sourceText">Text origin</param>
        /// <returns>existing item or null</returns>
        public LocalTranslationCacheItem TryGetValue(string sourceLanguage, string destLanguage, string sourceText)
        {
            foreach (LocalTranslationCacheItem item in this)
            {
                if (item.SourceLanguage.Equals(sourceLanguage, StringComparison.InvariantCultureIgnoreCase) &&
                     item.DestLanguage.Equals(destLanguage, StringComparison.InvariantCultureIgnoreCase) &&
                     item.SourceText.Equals(sourceText, StringComparison.InvariantCultureIgnoreCase))
                {
                    return item;
                }
            }
            return null;
        }
    }

    /// <summary>
    /// local cache value
    /// </summary>
    public class LocalTranslationCacheItem
    {
        /// <summary>
        /// create instance of the class
        /// </summary>
        /// <param name="sourceLanguage">Source Language</param>
        /// <param name="destLanguage">Target Language</param>
        /// <param name="sourceText">Origin text</param>
        /// <param name="translationText">Translated text</param>
        internal LocalTranslationCacheItem(string sourceLanguage, string destLanguage, string sourceText, string translationText)
        {
            SourceLanguage = sourceLanguage;
            DestLanguage = destLanguage;
            SourceText = sourceText;
            TranslationText = translationText;
        }

        /// <summary>
        /// Soruce Language
        /// </summary>
        public string SourceLanguage { get; private set; }
        
        /// <summary>
        /// Target Language
        /// </summary>
        public string DestLanguage { get; private set; }

        /// <summary>
        /// Origin Text
        /// </summary>
        public string SourceText { get; private set; }
        
        /// <summary>
        /// Translated Text
        /// </summary>
        public string TranslationText { get; private set; }
    }
}
