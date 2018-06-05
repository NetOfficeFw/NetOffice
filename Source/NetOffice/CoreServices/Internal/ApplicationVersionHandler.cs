using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Stores all registered version providers
    /// </summary>
    internal class ApplicationVersionHandler : IEnumerable<IApplicationVersionProvider>
    {
        private object _thisLock = new object();
        private List<IApplicationVersionProvider> _versionProviders = new List<IApplicationVersionProvider>();

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">parent core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal ApplicationVersionHandler(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
        }

        /// <summary>
        /// Parent Core
        /// </summary>
        internal Core Parent { get; private set; }

        /// <summary>
        /// All loaded application version providers
        /// </summary>
        internal IEnumerable<IApplicationVersionProvider> ApplicationVersionProviders
        {
            get
            {
                return _versionProviders.ToArray();
            }
        }

        /// <summary>
        /// Adds an provider to the instance
        /// </summary>
        /// <param name="provider">target provider as any</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal void Add(IApplicationVersionProvider provider)
        {
            if (null == provider)
                throw new ArgumentNullException("provider");
            lock (_thisLock)
            {
                _versionProviders.Add(provider);
            }            
        }

        /// <summary>
        /// Removes a provider from the instance
        /// </summary>
        /// <param name="provider">target provider as any</param>
        /// <returns>true if removed, otherwise false</returns>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal bool Remove(IApplicationVersionProvider provider)
        {
            if (null == provider)
                throw new ArgumentNullException("provider");
            lock (_thisLock)
            {
                return _versionProviders.Remove(provider);
            }
        }

        /// <summary>
        /// Removes all version providers from the instance
        /// </summary>
        internal void Clear()
        {
            lock (this)
            {
                _versionProviders.Clear();
            }
        }

        /// <summary>
        /// Returns application name and version
        /// </summary>
        /// <param name="componentName">component to look into</param>
        /// <returns>application version or empty</returns>
        internal string GetApplicationVersion(string componentName)
        {
            if (String.IsNullOrWhiteSpace(componentName))
                return String.Empty;
            lock (_thisLock)
            {
                var provider = _versionProviders.FirstOrDefault();
                if (null != provider)
                {
                    if (!provider.VersionRequested)
                        provider.TryRequestVersion();

                    if (null != provider.Version)
                        return String.Format("{0} {1}", provider.Name, provider.Version);
                    else
                        return String.Empty;
                }
                else
                    return String.Empty;
            }
        }

        /// <summary>
        /// Register an application version provider by its component name.
        /// The method does nothing if a version provider for a component with given name already exists.
        /// </summary>
        /// <param name="versionProvider">version provider as any</param>
        /// <returns>true if provider registerd, otherwise false</returns>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal bool RegisterApplicationVersionProvider(IApplicationVersionProvider versionProvider)
        {
            if (null == versionProvider)
                throw new ArgumentNullException("versionProvider");

            lock (_thisLock)
            {
                if (!_versionProviders.Any(e => e.ComponentName == versionProvider.ComponentName))
                {
                    if (Parent.Settings.ForceApplicationVersionProviders && false == versionProvider.VersionRequested)
                        versionProvider.TryRequestVersion();
                    _versionProviders.Add(versionProvider);
                    return true;
                }
                else
                    return false;
            }
        }

        /// <summary>
        /// Removes an application version provider
        /// </summary>
        /// <param name="versionProvider">version provider as any</param>
        /// <returns>true if removed, otherwise false</returns>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal bool UnregisterApplicationVersionProvider(IApplicationVersionProvider versionProvider)
        {
            if (null == versionProvider)
                throw new ArgumentNullException("versionProvider");

            lock (_thisLock)
            {
                return _versionProviders.Remove(versionProvider);
            }
        }

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>a System.Collections.Generic.IEnumerator`1 that can be used to iterate through the collection</returns>
        internal IEnumerable<IApplicationVersionProvider> ThreadSafeEnumerable()
        {
            lock (_thisLock)
            {
                return _versionProviders.ToArray();
            }
        }

        /// <summary>
        ///  Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>a System.Collections.Generic.IEnumerator`1 that can be used to iterate through the collection</returns>
        public IEnumerator<IApplicationVersionProvider> GetEnumerator()
        {
            return _versionProviders.GetEnumerator();
        }

        /// <summary>
        ///  Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>a System.Collections.Generic.IEnumerator`1 that can be used to iterate through the collection</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _versionProviders.GetEnumerator();
        }
    }
}
