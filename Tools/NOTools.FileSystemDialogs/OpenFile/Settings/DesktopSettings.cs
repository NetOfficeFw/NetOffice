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
        internal DesktopSettings(DefaultSettings defaultSettings, PropertyChangedEventHandler eventHandler = null) : base(defaultSettings, eventHandler)
        {
            
        }

        public override string ToString()
        {
            return "Desktop";
        }
    }
}
