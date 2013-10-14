using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.ConsoleMonitor
{
    /// <summary>
    /// Display the Admint Hint
    /// </summary>
    public partial class FormAdminMessage : Form
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public FormAdminMessage()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Just show the message
        /// </summary>
        /// <param name="parent">parent (main) window</param>
        public static void ShowAdminMessage(IWin32Window parent)
        {
            if (null == parent)
                throw new ArgumentNullException();
            FormAdminMessage window = new FormAdminMessage();
            window.ShowDialog(parent);
        }
    }
}
