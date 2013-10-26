using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.DeveloperToolbox
{
    /// <summary>
    /// define error categories
    /// </summary>
    public enum ErrorCategory
    {
        /// <summary>
        /// the error is non critical
        /// </summary>
        NonCritical = 0,

        /// <summary>
        /// the error is an critical/unexpected error
        /// </summary>
        Critical = 1,

        /// <summary>
        /// the error is a sudden death error. the program has to terminate immediately
        /// </summary>
        Penalty = 2
    }
}
