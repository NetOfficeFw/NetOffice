using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Informations
{
    /// <summary>
    /// Represents a diagnostics subset
    /// </summary>
    public class DiagnosticPair
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="type">name/type</param>
        /// <param name="value">value for the diagnostics</param>
        public DiagnosticPair(string type, string value)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            Type = type;
            Value = value;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name/Type of Diagnostics
        /// </summary>
        public string Type { get; private set; }

        /// <summary>
        /// Value (Can be null)
        /// </summary>
        public string Value { get; private set; }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("DiagnosticPair {0}:{1}", Type, Value);
        }

        #endregion
    }
}
