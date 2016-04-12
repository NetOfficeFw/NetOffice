using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.DeveloperToolbox.Controls.Hex
{
    internal static class Util
    {
        /// <summary>
        /// Gets true, if we are in design mode of Visual Studio
        /// </summary>
        /// <remarks>
        /// In Visual Studio 2008 SP1 the designer is crashing sometimes on windows forms. 
        /// The DesignMode property of Control class is buggy and cannot be used, so use our own implementation instead.
        /// </remarks>
        public static bool DesignMode
        {
            get
            {
                return Program.IsDesign;
            }
        }
    }
}
