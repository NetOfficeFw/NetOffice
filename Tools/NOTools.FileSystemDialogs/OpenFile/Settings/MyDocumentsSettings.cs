using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// All MyDocument settings for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class MyDocumentsSettings : DefaultableSettings
    {
        public MyDocumentsSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {
          
        }

        #region Overrides

        /// <summary>
        /// Returns a System.String that represence the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return "MyDocuments";
        }

        #endregion
    }
}
