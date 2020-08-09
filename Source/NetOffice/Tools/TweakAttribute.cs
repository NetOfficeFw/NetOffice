using System;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.Tools
{
    /// <summary>
    /// Activate tweaks for <see cref="COMAddinBase"/> derived classes.
    /// You can add various values in the Microsoft Office Add-ins registry location
    /// to customize NetOffice diagnostic/log settings at runtime. This can be very helpful in troubleshooting.
    /// See tweaks overview here: http://netoffice.codeplex.com/wikipage?title=Tweaks_EN
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple= false)]
    public class TweakAttribute : System.Attribute
    {
        /// <summary>
        /// Enable or disable possible tweaking is possible
        /// </summary>
        public readonly bool Enabled;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="enabled">Enable or disable possible tweaking is possible</param>
        public TweakAttribute(bool enabled)
        {
            Enabled = enabled;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public TweakAttribute()
        {
            Enabled = true;
        }
    }
}
