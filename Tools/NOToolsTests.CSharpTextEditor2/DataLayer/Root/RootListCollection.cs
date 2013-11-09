using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Collection for RootList instances
    /// </summary>
    public class RootListCollection : IEnumerable<RootList>
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public RootListCollection()
        {
            List = new List<RootList>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Inner list instance
        /// </summary>
        private List<RootList> List { get; set; }

        /// <summary>
        /// Returns a sepcific root list instance
        /// </summary>
        /// <param name="tableName">target name of the table</param>
        /// <returns>RootList instance</returns>
        public RootList this[string tableName]
        {
            get
            {
                foreach (var item in List)
                    if (item.Name.Equals(tableName, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                throw new ArgumentOutOfRangeException(tableName);
            }
        }

        /// <summary>
        /// Returns the count auf root list instances
        /// </summary>
        public int Count
        {
            get 
            {
                return List.Count;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add a new root list to the collection instance
        /// </summary>
        /// <param name="name">unique name of the new root list</param>
        /// <returns>new created instance</returns>
        public RootList Add(string name)
        {
            if (Contains(name))
                throw new InvalidOperationException(String.Format("A list {0} already exists in the collection",name));

            RootList list = new RootList(name);
            List.Add(list);

            return list;
        }

        /// <summary>
        /// Remove a root list instance
        /// </summary>
        /// <param name="name">name of the target list instance</param>
        public void Remove(string name)
        {
            RootList list = TryGet(name);
            if(null != list)
                List.Remove(list);
            else
                throw new ArgumentOutOfRangeException(name);
        }

        /// <summary>
        /// Remove a root list instance from the collection instance
        /// </summary>
        /// <param name="list">target list instance</param>
        public void Remove(RootList list)
        {
            List.Remove(list);
        }

        /// <summary>
        /// Returns info the collection includes a root list instance with specific name
        /// </summary>
        /// <param name="name">target name</param>
        /// <returns>true if exists, otherwise false</returns>
        public bool Contains(string name)
        {
            foreach (RootList item in this)
            {
                if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Try to get an root list instance with specific name
        /// </summary>
        /// <param name="name">name of the target root list</param>
        /// <returns>root list instance or null</returns>
        public RootList TryGet(string name)
        {
            foreach (RootList item in this)
            {
                if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    return item;
            }
            return null;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// RootList Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator<RootList> GetEnumerator()
        {
            return List.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return List.GetEnumerator();
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {
            return String.Format("{0} Items", Count);
        }

        #endregion
    }
}
