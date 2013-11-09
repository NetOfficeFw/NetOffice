using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Represents a top-level root table row
    /// </summary>
    public class RootItem : DataItem
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class.
        /// Stub ctor to create a local new item 
        /// </summary>
        public RootItem()
        { 
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="dataNode">xml node with data</param>
        public RootItem(XElement dataNode)
        {
            foreach (var item in dataNode.Attributes())
                Properties.Add(item.Name.LocalName, item.Value);

        }

        #endregion
    }
}
