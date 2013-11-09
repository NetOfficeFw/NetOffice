using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// List for AccessContextList instances
    /// </summary>
    internal class AccessContextListCollection : List<AccessContextList>
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">associated context</param>
        internal AccessContextListCollection(AccessContext parent)
        {
            Parent = parent;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Associated context
        /// </summary>
        internal AccessContext Parent { get; private set; }

        /// <summary>
        /// Returns a AccessContextList instance
        /// </summary>
        /// <param name="name">unique name of the list</param>
        /// <returns>AccessContextList instance</returns>
        public AccessContextList this[string name]
        {
            get
            {
                foreach (AccessContextList item in this)
                {
                    if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
                throw new ArgumentOutOfRangeException(name);
            }
        }

        #endregion
    }
}
