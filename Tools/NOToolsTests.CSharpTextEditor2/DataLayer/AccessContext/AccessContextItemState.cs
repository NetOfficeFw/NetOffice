using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Indicates the state of an AccessContext Item instance
    /// </summary>
    public enum AccessContextItemState
    {
        /// <summary>
        /// Item is local new created
        /// </summary>
        ItemIsNew = 0,
        
        /// <summary>
        /// Item is from database and contains no local uncomitted changes
        /// </summary>
        ItemIsNormal = 1,
        
        /// <summary>
        /// Item is from database and contains local uncomitted changes
        /// </summary>
        ItemIsLocalChanged = 2,
        
        /// <summary>
        /// Item is local deleted but the delete is not commited
        /// </summary>
        ItemIsLocalDeleted = 3,
        
        /// <summary>
        /// Item is completly deleted
        /// </summary>
        ItemIsDeleted = 4
    }
}
