using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Core Cache Holder
    /// </summary>
    internal class CoreCache : ICoreCache
    {
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal CoreCache(Core parent)
        {
            if (null == parent)
                throw new ArgumentNullException("parent");
            Parent = parent;
            EntitiesListCache = new Dictionary<string, Dictionary<string, string>>();
            TypeComponentIdCache = new Dictionary<Guid, Guid>();
            TypeCache = new TypeDictionary();
            VersionProviders = new ApplicationVersionHandler(Parent);
        }

        #endregion

        #region ICoreCache

        public Core Parent { get; private set; }

        public KeyValuePair<Guid, Guid>[] GetTypeComponentIdCache()
        {
            return TypeComponentIdCache.ToArray();
        }

        public IEnumerable<TypeInformation> GetTypeCache()
        {
            return TypeCache.ToEnumerable();
        }

        #endregion

        #region Methods

        /// <summary>
        /// ICOMObjectAvaility Cache
        /// </summary>
        internal Dictionary<string, Dictionary<string, string>> EntitiesListCache { get; private set; }

        /// <summary>
        /// Cache as Type ID (COM) => ParentLibrary(COM Component) ID 
        /// </summary>
        internal Dictionary<Guid, Guid> TypeComponentIdCache { get; private set; }

        /// <summary>
        /// Proxy,Contract,Implementation Type Cache
        /// </summary>
        internal TypeDictionary TypeCache { get; private set; }

        /// <summary>
        /// Registered Version Providers
        /// </summary>
        internal ApplicationVersionHandler VersionProviders { get; private set; }

        internal void Clear()
        {
            EntitiesListCache.Clear();
            TypeComponentIdCache.Clear();
            TypeCache.Clear();
            VersionProviders.Clear();
        }

        #endregion
    }
}
