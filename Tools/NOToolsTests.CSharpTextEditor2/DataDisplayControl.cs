using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NOToolsTests.CSharpTextEditor2.DataLayer;

namespace NOToolsTests.CSharpTextEditor2
{
    public partial class DataDisplayControl : UserControl, IDataDisplayControl
    {
        public DataDisplayControl()
        {
            InitializeComponent();
        }

        public void OnConnect(IDataHost parent, string tableName)
        {
            AccessContext context = parent.Tables[tableName].Add(new Guid().ToString());
            dataGridView1.DataSource = context;
        }

        public void OnShow(IDataHost parent)
        {
           
        }

        public void OnUnload(IDataHost parent)
        {
            
        }
    }
}
