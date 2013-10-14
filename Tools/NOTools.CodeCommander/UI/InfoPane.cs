using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.OfficeApi.Tools;
using NOTools.DeveloperAddin.Logic;

namespace NOTools.DeveloperAddin.UI
{
    public partial class InfoPane : UserControl
    {
        public InfoPane()
        {
            InitializeComponent();
        }

        public void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {

        }

        public void OnDisconnection()
        {

        }

        private void linkLabelInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Process.Start(linkLabelInfo.Tag as string);
            }
            catch
            {
                DialogBox.ShowSimpleError("Unable to start the NetOffice addin page.");
            }
        }

        private void checkBoxSaveSettings_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
