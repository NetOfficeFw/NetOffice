using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NOToolsTests.CSharpTextEditor2.DataLayer;

namespace NOToolsTests.CSharpTextEditor2
{
    public interface IDataDisplayControl
    {
        /// <summary>
        /// Called from host application after startup
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="tableName"></param>
        void OnConnect(IDataHost parent, string tableName);

        /// <summary>
        /// Called from host applicaton when user switch to component
        /// </summary>
        /// <param name="parent"></param>
        void OnShow(IDataHost parent);

        /// <summary>
        /// Called from host applicaton before shutdown
        /// </summary>
        /// <param name="parent"></param>
        void OnUnload(IDataHost parent);

        /// <summary>
        /// Local Context for the component
        /// </summary>
        AccessContext Context { get; }
    }
}
