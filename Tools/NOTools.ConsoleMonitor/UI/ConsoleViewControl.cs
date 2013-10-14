using System;
using System.Collections;
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
    /// Represents a Console
    /// </summary>
    public partial class ConsoleViewControl : UserControl, IApplicationControl
    {
        #region Ctor

        public ConsoleViewControl()
        {
            InitializeComponent();
            Items = new ConsoleViewItemCollection();
        }

        public ConsoleViewControl(IApplicationHost host, string consoleName)
        {
            InitializeComponent();
            Host = host;
            Items = new ConsoleViewItemCollection();
            ControlName = consoleName;
        }

        #endregion

        #region Events

        public event EventHandler CloseClick;

        private void RaiseCloseClick()
        {
            if (null != CloseClick && !String.IsNullOrWhiteSpace(ControlName))
                CloseClick(this, new EventArgs());
        }

        #endregion

        #region Properties

        public string ControlName { get; private set; }

        public IApplicationHost Host { get; internal set; }

        public ConsoleViewStyle ViewStyle { get; set; }

        public bool ShowCloseButton
        {
            get
            {
                return buttonCloseConsole.Visible;
            }
            set
            {
                buttonCloseConsole.Visible = value;
            }
        }

        internal ConsoleViewItemCollection Items { get; private set; }

        #endregion

        #region IApplicationControl

        public void UpdateDisplayContent(bool showTime, bool showMachine, bool showAppDomain)
        {
            TextBoxConsole.Text = Items.CreateText(ViewStyle, showTime, showMachine, showAppDomain);          
        }

        public void Clear()
        {
            Items.Clear();
            TextBoxConsole.Clear();
        }

        public string AddNewMessage(string notifyTime, string consoleName, string machineName, string appDomainFriendlyName, string parentEntryID, string message, bool showTime, bool showMachine, bool showAppDomain)
        {
            string newEntryID = Items.AddNew(machineName, appDomainFriendlyName, notifyTime, message, parentEntryID);
            if(Host.IsCurrentlyVisible(this))
                UpdateDisplayContent(showTime, showMachine, showAppDomain);
            return newEntryID;
        }

        #endregion

        #region Trigger
          
        private void radioButtonViewStyle_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonPlain.Checked)
                ViewStyle = ConsoleViewStyle.Plain;
            else if (radioButtonPlainReverse.Checked)
                ViewStyle = ConsoleViewStyle.PlainReverse;
            else if (radioButtonHierarchy.Checked)
                ViewStyle = ConsoleViewStyle.Hierarchy;
            UpdateDisplayContent(Host.ShowTime,Host.ShowMachine, Host.ShowAppDomain);
        }

        private void buttonCloseConsole_Click(object sender, EventArgs e)
        {
            RaiseCloseClick();
        }

        #endregion
    }
}
