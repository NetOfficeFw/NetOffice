using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace NOTools.FileSystemDialogs
{    
    /// <summary>
    /// All desktop settings for OpenFilePanel.cs
    /// </summary>
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class DesktopSettings : DefaultableSettings
    {
        #region Ctor

        internal DesktopSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {            
        }
        
        #endregion

        #region Override

        public override string ToString()
        {
            return "Desktop";
        }

        #endregion
    }
}
