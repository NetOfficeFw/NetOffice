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
        
        public AccessContext Context { get; private set; }

        public void OnConnect(IDataHost parent, string tableName)
        {
            Context = parent.Local.Add(Guid.NewGuid().ToString());
            dataGridView1.DataSource = Context[tableName];
        }

        public void OnShow(IDataHost parent)
        {
           
        }

        public void OnUnload(IDataHost parent)
        {
            parent.Local.Remove(Context);
            Context = null;
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void buttonCancelChanges_Click(object sender, EventArgs e)
        {

        }
    }
}
