using System;

namespace NetOffice.Diagnostics
{
    /// <summary>
    /// Diagnostics Data as Name/Value
    /// </summary>
    public class DiagnosticItem
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="name">name as any</param>
        /// <param name="value">value as any</param>
        public DiagnosticItem(string name, string value)
        {
            Name = name;
            Value = value;
        }

        /// <summary>
        /// Information Name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Information Value
        /// </summary>
        public string Value { get; private set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("{0}:{1}", Name, Value);
        }
    }
}
