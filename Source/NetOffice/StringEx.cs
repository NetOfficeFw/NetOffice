using System;

namespace NetOffice
{
    internal static class StringEx
    {
        /// <summary>
        /// Returns a value indicating whether a specified substring occurs within this string.
        /// This method uses case insensitive comparison.</summary>
        /// <param name="text">The original string.</param>
        /// <param name="value">The string to seek.</param>
        internal static bool ContainsIgnoreCase(this string text, string value)
        {
            return text?.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
