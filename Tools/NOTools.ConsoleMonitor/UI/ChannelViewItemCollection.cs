using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// 
    /// </summary>
    internal class ChannelViewItemCollection : List<ChannelViewItem>
    {
        #region Fields

        private static string _emptyMachineDefault = "<Unkown>";
        private static string _emptyAppDomainDefault = "<Undefined>";
        private static int _stringBuilderDefaultSize = 2048;

        private int _idCounter = 0;
        
        #endregion

        public new void Clear()
        {
            _idCounter = 0;
            base.Clear();
        }

        /// <summary>
        /// Create a System.String to display the items
        /// </summary>
        /// <param name="showTime">show sender time</param>
        /// <param name="showMachine">show sender machine</param>
        /// <param name="showAppDomain">show sender appDomain</param>
        /// <returns>String to display</returns>
        public string CreateText(bool showTime, bool showMachine, bool showAppDomain)
        {
            StringBuilder sb = new StringBuilder(_stringBuilderDefaultSize);

            foreach (var item in this)
            {
                string time = true == showTime ? "Time:" + item.Time + " " : "";
                string machine = true == showMachine ? "Machine:" + item.Machine + " " : "";
                string appDomain = true == showAppDomain ? "AppDomain:" + item.AppDomain + " " : "";
                sb.Append(String.Format("Channel:[{0}]{1}{2}{3}Value:{4}{5}", item.Channel, time, machine, appDomain, item.Text, Environment.NewLine));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Add a new display item
        /// </summary>
        /// unique name of the display channel
        /// <param name="machineName">name of the sender machine</param>
        /// <param name="appDomain">name of the sender appdomain</param>
        /// <param name="time">sender time</param>
        /// <param name="text">display text</param>
        /// <param name="parentID">parent display item id</param>
        /// <returns>the new created id for the item</returns>
        public string AddNew(string channelName, string machineName, string appDomain, string time, string text, string parentID = null)
        {
            string machineString = String.IsNullOrWhiteSpace(machineName) ? _emptyMachineDefault : machineName;
            string appDomainString = String.IsNullOrWhiteSpace(appDomain) ? _emptyAppDomainDefault : appDomain;

            foreach (var item in this)
            {
                if (item.Channel == channelName)
                {
                    item.Machine = machineString;
                    item.AppDomain = appDomainString;
                    item.Time = time;
                    item.Text = text;
                    return item.ID;
                }
            }

            _idCounter++;
            ChannelViewItem newItem = new ChannelViewItem(channelName, _idCounter.ToString(), machineString, appDomainString, time, text);
            Add(newItem);
            return _idCounter.ToString();
        }
    }
}
