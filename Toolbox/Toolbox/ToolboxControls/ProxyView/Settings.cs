using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox.ToolboxControls.ProxyView
{
    public static class Settings
    {
        static Settings()
        {
            ShowAllAccessible = true;
            ShowDetails = true;
            RefreshInterval = 2000;
        }

        public static bool ShowAllAccessible {get;set;}

        public static int RefreshInterval { get; set; }

        public static bool ShowDetails { get; set; }
    }
}