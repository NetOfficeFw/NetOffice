using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    internal class ChannelViewItem
    {
                /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// unique name of the display channel
        /// <param name="id">unique ID of the item instance</param>
        /// <param name="machine">name of the sender machine</param>
        /// <param name="appDomain">appDomain name of the sender process</param>
        /// <param name="time">sender time</param>
        /// <param name="text">display text</param>
        internal ChannelViewItem(string channelName, string id, string machine, string appDomain, string time, string text)
        {
            Channel = channelName;
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
        /// Name of the display channel
        /// </summary>
        public string Channel { get; internal set; }

        /// <summary>
        /// Sender time
        /// </summary>
        public string Time { get; internal set; }

        /// <summary>
        /// Name of the sender machine
        /// </summary>
        public string Machine { get; internal set; }

        /// <summary>
        /// AppDomain name of the sender process
        /// </summary>
        public string AppDomain { get; internal set; }

        /// <summary>
        /// Display text
        /// </summary>
        public string Text { get; internal set; }

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
