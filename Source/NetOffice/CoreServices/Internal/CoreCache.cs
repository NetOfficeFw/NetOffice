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
        #region Fields

        private TypeDictionary _typeCache;
        private Dictionary<Guid, Guid> _typeComponentIdCache;

        #endregion

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
            _typeCache = new TypeDictionary();
            _typeComponentIdCache = new Dictionary<Guid, Guid>();
            VersionProviders = new ApplicationVersionHandler(Parent);
        }

        #endregion

        #region Properties

        /// <summary>
        /// ICOMObjectAvaility Cache
        /// </summary>
        internal Dictionary<string, Dictionary<string, string>> EntitiesListCache { get; private set; }

        /// <summary>
        /// Cache as Type ID (COM) => ParentLibrary(COM Component) ID 
        /// </summary>
        internal Dictionary<Guid, Guid> TypeComponentIdCache
        {
            get
            {
                Parent.CheckInitialize();
                return _typeComponentIdCache;
            }
        }

        /// <summary>
        /// Proxy,Contract,Implementation Type Cache
        /// </summary>
        internal TypeDictionary TypeCache
        {
            get
            {
                Parent.CheckInitialize();
                return _typeCache;
            }
        }

        /// <summary>
        /// Registered Version Providers
        /// </summary>
        internal ApplicationVersionHandler VersionProviders { get; private set; }

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
