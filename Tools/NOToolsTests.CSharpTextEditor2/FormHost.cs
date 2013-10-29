using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOToolsTests.CSharpTextEditor2
{
    public partial class FormHost : Form , IDataHost
    {
        public FormHost()
        {
            InitializeComponent();
            dataDisplayControl1.OnConnect(this, "Persons");
        }

        public DataLayer.RootListDefinitionCollection Tables
        {
            get
            {
                if (null == _tables)
                    _tables = new DataLayer.RootListDefinitionCollection(new string[] { "Persons" });
                return _tables;
            }
        }
        private DataLayer.RootListDefinitionCollection _tables;
    }
}
