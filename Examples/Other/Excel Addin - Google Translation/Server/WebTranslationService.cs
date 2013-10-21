using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Sample.Server
{
    public delegate void TranslationEventHandler(TranslateOperationResult result);

    /// <summary>
    /// Offers Google Translation Functionality
    /// </summary>
    public class WebTranslationService : MarshalByRefObject
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public WebTranslationService()
        {
            Cache = new TranslationCache();
        }
        #endregion

        #region Properties
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
                    foreach (var item in GoogleTranslator.LanguageModeMap)
                        list.Add(item.Key);
                    _availableTranslations = list.ToArray();
                }

                return _availableTranslations;
            }
        }
        private static string[] _availableTranslations;

        #endregion

        #region Methods

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
                LocalTranslationCacheItem cacheItem = Cache.TryGetValue(sourceLanguage, destLanguage, text);
                if (null != cacheItem)
                {
                    TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Sucseed, text, cacheItem.TranslationText, null, true);
                    RaiseOnTranslation(result);
                    return cacheItem.TranslationText;
                }
                else
                {
                    GoogleTranslator translator = new GoogleTranslator(sourceLanguage, destLanguage, text);
                    translator.Translate();
                    TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Sucseed, text, translator.Translation, null);
                    Cache.Add(sourceLanguage, destLanguage, text, translator.Translation);
                    RaiseOnTranslation(result);
                    return translator.Translation;    
                }
            }
            catch (Exception exception)
            {
                TranslateOperationResult result = new TranslateOperationResult(TranslateOperationState.Error, text, null, exception);
                RaiseOnTranslation(result);
                throw exception;
            }
        }

        #endregion

        #region Events

        private void RaiseOnTranslation(TranslateOperationResult result)
        {
            foreach (var item in Reapters)
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

        public static DataEventRepeators Reapters
        {
            get
            {
                if (reapters == null)
                {
                    reapters = new DataEventRepeators();
                }
                return WebTranslationService.reapters;
            }
        }
        private static DataEventRepeators reapters;

        public void AddEventRepeater(DataEventRepeator repeater)
        {
            Reapters.Add(repeater);
        }
    }  
}
