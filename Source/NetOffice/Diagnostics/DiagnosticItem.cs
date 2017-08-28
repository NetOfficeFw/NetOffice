using System;

namespace NetOffice.Diagnostics
{
    /// <summary>
    /// Data item
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
    }
}
