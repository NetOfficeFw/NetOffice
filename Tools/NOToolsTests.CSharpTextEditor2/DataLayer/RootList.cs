using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    public class RootList: DataList<RootItem>
    {
        public RootList(string listName) : base(listName)
        {
            LoadFromDatabase();
        }

        public override void LoadFromDatabase()
        {
            Clear();
            XDocument dataDocument = XDocument.Parse(ReadXmlString(ListName + ".xml"));
            foreach (var item in (dataDocument.FirstNode as XElement).Elements("Customer"))
                Add(new RootItem(item));
        }
    }
}
