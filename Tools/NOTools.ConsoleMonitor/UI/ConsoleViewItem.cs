using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Display view item
    /// </summary>
    internal class ConsoleViewItem
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="id">unique ID of the item instance</param>
        /// <param name="machine">name of the sender machine</param>
        /// <param name="appDomain">appDomain name of the sender process</param>
        /// <param name="time">sender time</param>
        /// <param name="text">display text</param>
        internal ConsoleViewItem(string id, string machine, string appDomain, string time, string text)
        {
            ID = id;
            Machine = machine;
            AppDomain = appDomain;
            Time = time;
            Text = text;
            Items = new List<ConsoleViewItem>();
        }

        /// <summary>
        /// Unique ID of the item instance
        /// </summary>
        public string ID { get; private set; }

        /// <summary>
        /// Sender time
        /// </summary>
        public string Time { get; private set; }

        /// <summary>
        /// Name of the sender machine
        /// </summary>
        public string Machine { get; private set; }

        /// <summary>
        /// AppDomain name of the sender process
        /// </summary>
        public string AppDomain { get; private set; }

        /// <summary>
        /// Display text
        /// </summary>
        public string Text { get; private set; }

        /// <summary>
        /// Sub items of the item instance
        /// </summary>
        public List<ConsoleViewItem> Items { get; private set; }

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("[{0}]{1}", ID, String.IsNullOrWhiteSpace(Text) ? "<Empty>" : Text);
        }
    }
}
