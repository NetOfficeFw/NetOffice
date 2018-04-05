using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NetOffice.DeveloperToolbox.ToolboxControls.RegistryEditor
{
    public partial class CancelSearchControl : UserControl
    {
        public CancelSearchControl(Action cancelRequest)
        {
            if (null == cancelRequest)
                throw new ArgumentNullException("cancelRequest");
            InitializeComponent();
            CancelLinkLabel.LinkClicked += delegate
            {
                cancelRequest();
            };
        }
    }
}