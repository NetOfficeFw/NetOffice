using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NOTools.CSharpTextEditor
{
    internal class ListViewGacItemComparer : IComparer<ListViewItem>
    {
        internal ListViewGacItemComparer(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        private int ColumnIndex { get; set; }

        public int Compare(ListViewItem x, ListViewItem y)
        {
            return (String.Compare(x.SubItems[ColumnIndex].Text, y.SubItems[ColumnIndex].Text));
        }
    }
}
