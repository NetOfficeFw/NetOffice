using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetOffice.Loader;
using System.Reflection;
using NetOffice.Exceptions;

namespace NetOffice.CoreServices.Internal
{
    /// <summary>
    /// Core Factory Holder
    /// </summary>
    internal class CoreFactories : ICoreFactories
    {
        #region Fields

        private object _thisLock = new object();

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="parent">affected netoffice core</param>
        /// <exception cref="ArgumentNullException">argument is null</exception>
        internal CoreFactories(Core parent)
        {
            if (null == parent)
                throw new ArgumentException("parent");
            Parent = parent;
            FactoryAssemblies = new FactoryList();
            DependentAssemblies = new List<DependentAssembly>();
        }

        #endregion

        #region ICoreFactories

        /// <summary>
        /// Affected NetOffice Core
        /// </summary>
        public Core Parent { get; private set; }

        /// <summary>
        /// Loaded Factories
        /// </summary>
        public IEnumerable<ITypeFactory> Factories
        {
            get
            {
                lock (_thisLock)
                {
                    ITypeFactory[] result = new ITypeFactory[FactoryAssemblies.Count];
                    for (int i = 0; i < FactoryAssemblies.Count; i++)
                        result[i] = FactoryAssemblies[i];
                    return result;
                }
            }
        }

        /// <summary>
        /// Known dependent NetOffice assemblies
        /// </summary>
        public IEnumerable<DependentAssembly> Dependents
        {
            get
            {
                lock (_thisLock)
                {
                    DependentAssembly[] result = new DependentAssembly[DependentAssemblies.Count];
                    for (int i = 0; i < DependentAssemblies.Count; i++)
                        result[i] = DependentAssemblies[i];
                    return result;
                }
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Returns a sequence of currently loaded NetOffice API assemblies
        /// </summary>
        internal FactoryList FactoryAssemblies { get; private set; }

        /// <summary>
        /// Dependent assemblies analyzed by LoadAPIFactories method
        /// </summary>
        internal List<DependentAssembly> DependentAssemblies { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Removes all factory informations from the instance
        /// </summary>
        internal void Clear()
        {
            lock (_thisLock)
            {
                FactoryAssemblies.Clear();
                DependentAssemblies.Clear();
            }
        }

        /// <summary>
        /// Removes all dependent assembly informations from the instance
        /// </summary>
        internal void ClearDependentAssemblies()
        {
            lock (this)
            {
                DependentAssemblies.Clear();
            }
        }

        /// <summary>
        /// Adds an dependent assembly information to the instance
        /// </summary>
        /// <param name="name">assembly name</param>
        /// <param name="parentAssembly">parent assembly</param>
        internal void AddDependentAssembly(string name, Assembly parentAssembly)
        {
            DependentAssemblies.Add(new DependentAssembly(name, parentAssembly));
        }

        /// <summary>
        /// Determines a depenedent assembly is already registered
        /// </summary>
        /// <param name="name">name of the dependent assembly</param>
        /// <returns>true if registered, otherwise false</returns>
        internal bool ContainsDependentAssembly(string name)
        {
            return DependentAssemblies.Any(e => e.Name == name);
        }

        /// <summary>
        /// Analyze assemblies in current appdomain and connect all NetOffice API factories to the core runtime.
        /// </summary>
        internal void LoadAPIFactories()
        {
            ClearDependentAssemblies();
            Assembly[] assemblies = Parent.CoreDomain.GetAssemblies();
            foreach (Assembly itemAssembly in assemblies)
            {
                string assemblyName = itemAssembly.GetName().Name;
                if (Parent.KnownNetOfficeKeyTokens.ContainsNetOfficeAttribute(itemAssembly))
                {
                    string[] depends = RecieveAssemblyFactory(assemblyName, itemAssembly);
                    foreach (string depend in depends)
                    {
                        if (!ContainsDependentAssembly(depend))
                            AddDependentAssembly(depend, itemAssembly);
                    }
                }

                if (Parent.Settings.EnableDeepLoading)
                {
                    foreach (AssemblyName itemName in itemAssembly.GetReferencedAssemblies())
                    {
                        if (Parent.KnownNetOfficeKeyTokens.ContainsNetOfficePublicKeyToken(itemName))
                        {
                            Assembly deepAssembly = Parent.CoreDomain.Load(itemName);
                            if (null == deepAssembly)
                                continue;

                            string deepAssemblyName = itemName.Name;
                            string[] depends = RecieveAssemblyFactory(deepAssemblyName, deepAssembly);
                            foreach (string depend in depends)
                            {
                                if (!ContainsDependentAssembly(depend))
                                    AddDependentAssembly(depend, itemAssembly);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Analyze dependent assemblies and connect there NetOffice API factories to the core runtime
        /// </summary>
        internal void LoadDependentAPIFactories()
        {
            if (!Parent.Settings.EnableAdHocLoading)
                return;

            foreach (DependentAssembly dependAssembly in DependentAssemblies)
            {
                if (!FactoryAssemblies.Contains(dependAssembly.Name))
                {
                    string fileName = PathBuilder.BuildLocalPathFromDependentAssembly(dependAssembly);
                    if (System.IO.File.Exists(fileName))
                    {
                        try
                        {
                            Assembly asssembly = Parent.CoreDomain.Load(fileName);
                            if(null != asssembly)
                                RecieveAssemblyFactory(asssembly.GetName().Name, asssembly);
                        }
                        catch (Exception exception)
                        {
                            Parent.Console.WriteException(exception);
                        }
                    }
                    else
                    {
                        Parent.Console.WriteLine(string.Format("Assembly {0} not found.", fileName));
                    }
                }
            }
        }

        /// <summary>
        /// Recieve factory instance from assembly and add them to factory cache
        /// </summary>
        /// <param name="name">name of the assembly</param>
        /// <param name="assembly">assemmbly to recieve</param>
        /// <returns>array of dependend assemblies</returns>
        private string[] RecieveAssemblyFactory(string name, Assembly assembly)
        {
            if (false == Attributes.NetOfficeAssemblyAttribute.ContainsAttribute(assembly))
                return new string[0];

            ITypeFactory factoryInfo = FactoryAssemblies.FirstOrDefault(e => e.FactoryName == name);
            if (null == factoryInfo)
            {
                string targetTypeFactoryName = String.Format("NetOffice.{0}.{1}", name, "Tools.Expose.TypeFactory");

                List<string> dependAssemblies = new List<string>();
                Type factoryInfoType = assembly.GetType(targetTypeFactoryName, false);

                if (null == factoryInfoType)
                    throw new NetOfficeException(String.Format("Unable to find {0} type factory", name));

                factoryInfo = Activator.CreateInstance(factoryInfoType) as ITypeFactory;

                if (null == factoryInfo)
                    throw new FactoryException(String.Format("Unexpected {0} type factory. Assembly {0}", name, assembly));

                FactoryAssemblies.Add(factoryInfo);

                Parent.Console.WriteLine("NetOffice Core recieved factory:{0}:{1}", factoryInfo.Assembly.FullName, factoryInfo.Assembly.FullName);

                foreach (string itemDependency in factoryInfo.Dependencies)
                    dependAssemblies.Add(itemDependency);

                return dependAssemblies.ToArray();
            }
            else
                return new string[0];
        }

        #endregion
    }
}
