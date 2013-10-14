using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.OfficeApi.Tools;
using NOTools.DeveloperAddin.Logic;

namespace NOTools.DeveloperAddin.UI
{
    public partial class CommandPane : UserControl
    {
        public CommandPane()
        {
            InitializeComponent();
        }

        public void OnConnection(COMObject application, NetOffice.OfficeApi._CustomTaskPane parentPane, object[] customArguments)
        {
            Addin parent = customArguments[0] as Addin;
            gridCommands.DataSource = parent.Commands;
        }

        public void OnDisconnection()
        {

        }

        private void buttonAddNew_Click(object sender, EventArgs e)
        {

        }

        private void buttonExecuteCommand_Click(object sender, EventArgs e)
        {

        }
    }
}
