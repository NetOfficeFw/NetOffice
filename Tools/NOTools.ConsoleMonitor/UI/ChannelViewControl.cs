using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Control to display incoming channel messages
    /// </summary>
    public partial class ChannelViewControl : UserControl, IApplicationControl
    {
        #region Ctor

        public ChannelViewControl()
        {
            InitializeComponent();
            Items = new ChannelViewItemCollection();
        }

        public ChannelViewControl(IApplicationHost host, string consoleName)
        {
            InitializeComponent();
            Host = host;
            Items = new ChannelViewItemCollection();
            ControlName = consoleName;
        }

        #endregion

        #region Properties

        internal ChannelViewItemCollection Items { get; private set; }

        #endregion

        #region IApplicationControl

        /// <summary>
        /// 
        /// </summary>
        public string ControlName { get; private set; }

        public IApplicationHost Host { get; internal set; }
        
        public void UpdateViewOptions(bool showTime, bool showMachine, bool showAppDomain)
        {
            colTime.Width = true == showTime ? 100 : 0;
            colMachine.Width = true == showMachine ? 100 : 0;
            colAppDomain.Width = true == showAppDomain ? 100 : 0;
            ChannelViewControl_Resize(this, new EventArgs());
        }  

        public void UpdateDisplayContent(bool showTime, bool showMachine, bool showAppDomain)
        {
            UpdateViewOptions(showTime, showMachine, showAppDomain);
            ListViewChannels.Items.Clear();
            int i = 0;
            foreach (var item in Items)
            {                
                ListViewItem newItem = ListViewChannels.Items.Add(item.Channel);
                newItem.SubItems.Add(item.Time);
                newItem.SubItems.Add(item.Machine);
                newItem.SubItems.Add(item.AppDomain);
                newItem.SubItems.Add(item.Text);

                if (i == 256)
                    return;
                i++;
            }
        }

        public void Clear()
        {
            Items.Clear();
            ListViewChannels.Items.Clear();
        }

        public string AddNewMessage(string notifyTime, string channelName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            string newEntryID = Items.AddNew(channelName, machineName, appDomainFriendlyName, notifyTime, message, parentEntryID);
            if (Host.IsCurrentlyVisible(this))
                UpdateDisplayContent(showTime, showMachine, showAppDomain);
            return newEntryID;
        }

        #endregion

        #region Trigger
        
        private void ChannelViewControl_Resize(object sender, EventArgs e)
        {
            colMessage.Width = ListViewChannels.Width - (colTime.Width + colChannel.Width  + colMachine.Width + colAppDomain.Width + 10);
        }

        #endregion
    }
}
