using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOToolsTests.CSharpTextEditor2.DataLayer
{
    /// <summary>
    /// Collection for AccessContext Instances
    /// </summary>
    public class AccessContextCollection : IEnumerable<AccessContext>
    {
        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="dataSources">list of all available root tables</param>
        public AccessContextCollection(RootListCollection dataSources)
        {
            DataSources = dataSources;
            List = new List<AccessContext>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// List of all available root tables
        /// </summary>
        internal RootListCollection DataSources { get; private set; }

        /// <summary>
        /// Inner AccessContext List
        /// </summary>
        private List<AccessContext> List { get; set; }

        /// <summary>
        /// Returns an AccessContext instance
        /// </summary>
        /// <param name="name">unique name of the target context</param>
        /// <returns>AccessContext instance</returns>
        public AccessContext this[string name]
        {
            get 
            {
                foreach (AccessContext item in this)
                {
                    if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                        return item;
                }
                throw new ArgumentOutOfRangeException(name);
            }
        }

        /// <summary>
        /// Count of all AccessContext instances
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
        /// Add a new AccessContext instance
        /// </summary>
        /// <param name="name">unique name of the new context</param>
        /// <returns>new created AccessContext instance</returns>
        public AccessContext Add(string name)
        {
            if (Contains(name))
                throw new InvalidOperationException(String.Format("A context {0} already exists in the collection", name));

            AccessContext context = new AccessContext(this, name);
            List.Add(context);

            return context;
        }

        /// <summary>
        /// Remove an AccessContext instance
        /// </summary>
        /// <param name="name">unique name the instance</param>
        public void Remove(string name)
        {
            AccessContext list = TryGet(name);
            if (null != list)
                List.Remove(list);
            else
                throw new ArgumentOutOfRangeException(name);
        }

        /// <summary>
        /// Remove an AccessContext instance
        /// </summary>
        /// <param name="context">target AccessContext instance</param>
        public void Remove(AccessContext context)
        {
            List.Remove(context);
        }

        /// <summary>
        /// Returns info the collection contains an AccessContext instance with specific name
        /// </summary>
        /// <param name="name">target name</param>
        /// <returns>true if the collection contains an instance with specified name</returns>
        public bool Contains(string name)
        {
            foreach (AccessContext item in this)
            {
                if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Try to find an instance with specific name in the collection
        /// </summary>
        /// <param name="name">target name</param>
        /// <returns>AccessContext or null</returns>
        public AccessContext TryGet(string name)
        {
            foreach (AccessContext item in this)
            {
                if (item.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase))
                    return item;
            }
            return null;
        }

        /// <summary>
        /// Update data in all access contexts from a secific proxy table.
        /// The method was called from an access context after commit local changes to synchronize local view data.
        /// </summary>
        /// <param name="instance">originator proxy table instance</param>
        internal void UpdateNotifyOtherListInstances(AccessContextList instance)
        {
            foreach (AccessContext context in this)
                foreach (AccessContextList list in context.Tables)
                    if (list.DataSource == instance.DataSource && list != instance)
                        list.UpdateFromOtherInstance(instance);
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// AccessContext Enumerator
        /// </summary>
        /// <returns>Enumerator Instance</returns>
        public IEnumerator<AccessContext> GetEnumerator()
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
