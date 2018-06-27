using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Collect all currently open instances from a core
    /// </summary>
    internal class CoreManagement : List<ICOMObject>, ICoreManagement
    {
        #region Fields

        private object _thisLock = new object();
        private static ICOMObject[] _emptyOwnerPath = new ICOMObject[0];

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal CoreManagement(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        #endregion

        #region ICOMObjectManagement

        /// <summary>
        /// Notify info the count of proxies there open are changed
        /// in case of notify comes from event trigger created proxy the call comes from other thread
        /// </summary>
        public event CountChangedHandler CountChanged;

        /// <summary>
        /// Occurs when a proxy has been added
        /// </summary>
        public event AddedHandler Added;

        /// <summary>
        ///  Occurs when a proxy has been removed
        /// </summary>
        public event RemovedHandler Removed;

        /// <summary>
        /// Occurs when all proxies has been removed
        /// </summary>
        public event ClearHandler Cleared;

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }

        /// <summary>
        /// Returns all root instances in COM proxy management
        /// </summary>
        /// <returns>Enumerable sequence of root instances</returns>
        public IEnumerable<ICOMObject> GetRootInstances()
        {
            List<ICOMObject> result = new List<ICOMObject>();

            try
            {
                lock (_thisLock)
                {
                    foreach (ICOMObject item in this)
                    {
                        if (null == item.ParentObject)
                            result.Add(item);
                    }
                }
            }
            catch (Exception throwedException)
            {
                Parent.Console.WriteException(throwedException);
            }

            return result;
        }

        /// <summary>
        /// Dispose all open objects
        /// </summary>
        public void DisposeAllInstances()
        {
            lock (_thisLock)
            {
                // NetOffice is appending new proxies so we free them in reverse order
                while (Count > 0)
                    this[Count - 1].Dispose();
                Cleared?.Invoke(Parent);
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Added event has one ore more active event listeners
        /// </summary>
        public bool HasAddedRecipients
        {
            get
            {
                return null != Added;
            }
        }

        /// <summary>
        /// Removed event has one ore more active event listeners
        /// </summary>
        public bool HasRemovedRecipients
        {
            get
            {
                return null != Removed;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add object to global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        internal void AddObjectToList(ICOMObject proxy)
        {
            try
            {
                lock (_thisLock)
                {
                    Add(proxy);

                    if (null != Added)
                    {
                        IEnumerable<ICOMObject> ownerPath = GetOwnerPath(proxy);
                        Added?.Invoke(Parent, ownerPath, proxy);
                    }
                }
                CountChanged?.Invoke(Parent, Count);
            }
            catch (Exception throwedException)
            {
                Parent.Console.WriteException(throwedException);
            }
        }

        /// <summary>
        /// Remove object from global list
        /// </summary>
        /// <param name="proxy">com wrapper instance</param>
        /// <param name="ownerPath">optional owner path</param>
        internal void RemoveObjectFromList(ICOMObject proxy, IEnumerable<ICOMObject> ownerPath)
        {
            try
            {
                bool removed = false;
                lock (_thisLock)
                {
                    removed = Remove(proxy);
                    if(removed)
                        Removed?.Invoke(Parent, ownerPath, proxy);
                }
                if (removed)
                    CountChanged?.Invoke(Parent, Count);
            }
            catch (Exception throwedException)
            {
                Parent.Console.WriteException(throwedException);
            }
        }

        /// <summary>
        /// Returns an array with full parent(s) path
        /// </summary>
        /// <param name="comObject">target com object</param>
        /// <returns>top down path sequence</returns>
        internal static IEnumerable<ICOMObject> GetOwnerPath(ICOMObject comObject)
        {
            if (null == comObject.ParentObject)
                return _emptyOwnerPath;

            ICOMObject parent = comObject.ParentObject;
            int parentCount = 0;
            while (null != parent)
            {
                parentCount++;
                parent = parent.ParentObject;
            }

            ICOMObject[] result = new ICOMObject[parentCount];
            parent = comObject.ParentObject;
            while (null != parent)
            {
                result[parentCount - 1] = parent;
                parentCount--;
                parent = parent.ParentObject;
            }

            return result;
        }

        #endregion
    }
}
