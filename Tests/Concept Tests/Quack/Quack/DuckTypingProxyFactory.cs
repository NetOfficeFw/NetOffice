using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Reflection.Emit;
using System.Threading;

namespace Quack
{
    public class DuckTypingProxyFactory
    {
        #region static members
        private static readonly IDictionary<Tuple<Type, Type>, Type> _typeCache = new Dictionary<Tuple<Type, Type>, Type>();
        private static readonly string _assemblyName = "SkyLinq.DuckTypingProxies.Generated";
        private static readonly AssemblyBuilder _assemblyBuilder;
        private static readonly ModuleBuilder _moduleBuilder;
        private static readonly ReaderWriterLockSlim _cacheLock = new ReaderWriterLockSlim();

        static DuckTypingProxyFactory()
        {
            _assemblyBuilder = CodeGenUtil.CreateAssemblyBuilder(_assemblyName);
            _moduleBuilder = CodeGenUtil.CreateModuleBuilder(_assemblyBuilder, _assemblyName);
        }

        private static void ValidateParams(Type typeOfIMyDuck, object otherDuck)
        {
            if (!typeOfIMyDuck.IsInterface)
                throw new ArgumentException("proxyInterfaceType must be a type of an interface.");

            if (otherDuck == null)
                throw new ArgumentNullException("Object to be wrapped cannot be null");
        }

        #endregion

        public TIMyDuck GenerateProxy<TIMyDuck>(object otherDuck) where TIMyDuck : class
        {
            Type typeOfIMyDuck = typeof(TIMyDuck);
            ValidateParams(typeOfIMyDuck, otherDuck);

            //If obj already implements TProxy, simply return it
            TIMyDuck o = otherDuck as TIMyDuck;
            if (o != null)
                return o;

            Type proxyType = null;
            Type typeOfOtherDuck = otherDuck.GetType();
            _cacheLock.EnterUpgradeableReadLock();
            try
            {
                if (!_typeCache.TryGetValue(new Tuple<Type, Type>(typeOfIMyDuck, typeOfOtherDuck), out proxyType))
                {
                    //Generate the proxyType here
                    if (!CanBeDuckTypedTo<TIMyDuck>(otherDuck))
                        throw new ArgumentException("Object cannot be duck typed by the interface.");

                    proxyType = GenerateProxyType(typeOfIMyDuck, otherDuck.GetType());

                    _cacheLock.EnterWriteLock();
                    try
                    {
                        _typeCache.Add(new Tuple<Type, Type>(typeOfIMyDuck, typeOfOtherDuck), proxyType);
                    }
                    finally
                    {
                        _cacheLock.ExitWriteLock();
                    }
                }
                return (TIMyDuck)Activator.CreateInstance(proxyType, otherDuck);
            }
            finally
            {
                _cacheLock.ExitUpgradeableReadLock();
            }
        }

        public bool CanBeDuckTypedTo<TIMyDuck>(object otherDuck)
        {
            Type typeOfIMyDuck = typeof(TIMyDuck);
            ValidateParams(typeOfIMyDuck, otherDuck);
            Type t = otherDuck.GetType();
            return typeOfIMyDuck.GetMembers().All(m =>
            {
                //Interface member can be either method or property
                if (m.MemberType == MemberTypes.Method)
                {
                    MethodInfo mi = (MethodInfo)m;
                    MethodInfo mi2 = t.GetMethod(m.Name, mi.GetParameters().Select(pi => pi.ParameterType).ToArray());
                    return (mi2 != null) && mi2.IsPublic && !mi2.IsAbstract && !mi2.IsStatic && mi.ReturnType == mi2.ReturnType;
                }
                else //Property
                {
                    PropertyInfo pi = (PropertyInfo)m;
                    PropertyInfo pi2 = t.GetProperty(m.Name, pi.PropertyType, pi.GetIndexParameters().Select(param => param.ParameterType).ToArray());
                    return pi2 != null;
                }
            }
            );
        }

        public Type GenerateProxyType(Type typeOfIMyDuck, Type typeOfOtherDuck)
        {
            TypeAttributes newAttributes = TypeAttributes.Public | TypeAttributes.Class;
            TypeBuilder typeBuilder = _moduleBuilder.DefineType(
                typeOfOtherDuck.Name + "_DuckTypingProxy" + Guid.NewGuid().ToString(), newAttributes);

            // Add interface implementation
            typeBuilder.AddInterfaceImplementation(typeOfIMyDuck);

            FieldBuilder targetField = typeBuilder.DefineField("target", typeOfOtherDuck, FieldAttributes.Private);

            foreach (MethodInfo mi in typeOfIMyDuck.GetMethods())
            {
                CodeGenUtil.CreateDelegateImplementation(typeBuilder, targetField, mi);
            }

            foreach (PropertyInfo pi in typeOfIMyDuck.GetProperties())
            {
                PropertyBuilder pb = typeBuilder.DefineProperty(
                    pi.Name,
                    pi.Attributes,
                    pi.PropertyType,
                    pi.GetIndexParameters().Select(param => param.ParameterType).ToArray()
                );
                MethodInfo getMi = pi.GetGetMethod();
                if (getMi != null)
                {
                    MethodBuilder getMethod = CodeGenUtil.CreateDelegateImplementation(typeBuilder, targetField, getMi);
                    pb.SetGetMethod(getMethod);
                }

                MethodInfo setMi = pi.GetSetMethod();
                if (setMi != null)
                {
                    MethodBuilder setMethod = CodeGenUtil.CreateDelegateImplementation(typeBuilder, targetField, setMi);
                    pb.SetSetMethod(setMethod);
                }
            }

            //Constructor
            ConstructorBuilder ctorBuilder = typeBuilder.DefineConstructor(
                MethodAttributes.Public | MethodAttributes.HideBySig | MethodAttributes.SpecialName | MethodAttributes.RTSpecialName,
                CallingConventions.HasThis,
                new Type[] { typeOfOtherDuck });
            ctorBuilder.DefineParameter(1, ParameterAttributes.None, "target");
            ILGenerator il = ctorBuilder.GetILGenerator();

            // Call base class constructor
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Call, CodeGenUtil.GetConstructorInfo(() => new object()));

            // Initialize the target field
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldarg_1);
            il.Emit(OpCodes.Stfld, targetField);
            il.Emit(OpCodes.Ret);

            Type result = typeBuilder.CreateType();
#if DEBUG
            _assemblyBuilder.Save(_assemblyName + ".dll");
#endif
            return result;
        }
    }
}
