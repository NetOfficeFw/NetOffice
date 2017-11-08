using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice
{
    /*
    Why so many overloads instead of "params object[]" or optional arguments ? 

    Because for "params" the Compiler generates an argument array for the client caller in MSIL client assembly
    and it looks like thats bigger than push each argument on stack until it is less than 8 arguments.

    In order to shrink the size of API assemblies as best as possible - we give 4 fixed argument overloads too.
    (API assemblies in 1.7.4.1 call fixed arguments overloads whenever its possible)
    */
    
    /// <summary>
    /// Provides top-off Core/Invoker get property services to shrink caller code in Api assemblies and give more refactoring possibilies
    /// </summary>
    public static class CorePropertyGetExtensions
    {
        #region Fields

        private static object[] _emptyParams = new object[0];

        #endregion

        #region ExecuteObjectPropertyGet
        
        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteObjectPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteObjectPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteObjectPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteObjectPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteObjectPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static object ExecuteObjectPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteObjectPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static object ExecuteObjectPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            return value.Invoker.PropertyGet(caller, name, validatedArgs);
        }

        #endregion

        #region ExecuteInt16PropertyGet

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteInt16PropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteInt16PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteInt16PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteInt16PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteInt16PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Int16 ExecuteInt16PropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteInt16PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static Int16 ExecuteInt16PropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToInt16(returnItem) : (short)0;
        }

        #endregion

        #region ExecuteInt32PropertyGet

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteInt32PropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteInt32PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteInt32PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteInt32PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteInt32PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static int ExecuteInt32PropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteInt32PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static int ExecuteInt32PropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToInt32(returnItem) : 0;
        }

        #endregion

        #region ExecuteInt64PropertyGet

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteInt64PropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteInt64PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteInt64PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteInt64PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteInt64PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Int64 ExecuteInt64PropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteInt64PropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments</param>
        internal static Int64 ExecuteInt64PropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToInt64(returnItem) : 0;
        }

        #endregion

        #region ExecuteUIntPtrPropertyGet

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteUIntPtrPropertyGetInternal(value, caller, name, args);     
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static UIntPtr ExecuteUIntPtrPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? (UIntPtr)returnItem : UIntPtr.Zero;
        }

        #endregion

        #region ExecuteFloatPropertyGet

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static float ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteFloatPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static float ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteFloatPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static float ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteFloatPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static float ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteFloatPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static float ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteFloatPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Single ExecuteFloatPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteFloatPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static Single ExecuteFloatPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToSingle(returnItem) : 0;
        }

        #endregion

        #region ExecuteDoublePropertyGet

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteDoublePropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteDoublePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteDoublePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteDoublePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteDoublePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static double ExecuteDoublePropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteDoublePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static double ExecuteDoublePropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToDouble(returnItem) : 0;
        }

        #endregion

        #region ExecuteSinglePropertyGet

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteSinglePropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteSinglePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteSinglePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteSinglePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteSinglePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Single ExecuteSinglePropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteSinglePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static Single ExecuteSinglePropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToSingle(returnItem) : 0;
        }

        #endregion

        #region ExecuteDateTimePropertyGet

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteDateTimePropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteDateTimePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteDateTimePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteDateTimePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteDateTimePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteDateTimePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static DateTime ExecuteDateTimePropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToDateTime(returnItem) : default(DateTime);
        }

        #endregion

        #region ExecuteBytePropertyGet

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteBytePropertyGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteBytePropertyGet(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteBytePropertyGet(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteBytePropertyGet(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteBytePropertyGet(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static byte ExecuteBytePropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
            return null != returnItem ? Convert.ToByte(returnItem) : default(byte);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static byte ExecuteBytePropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToByte(returnItem) : default(byte);
        }

        #endregion

        #region ExecuteBoolPropertyGet

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteBoolPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteBoolPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteBoolPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteBoolPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteBoolPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray"> arguments as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteBoolPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static bool ExecuteBoolPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToBoolean(returnItem) : false;
        }

        #endregion

        #region ExecuteStringPropertyGet

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteStringPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteStringPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteStringPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteStringPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteStringPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static string ExecuteStringPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteStringPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static string ExecuteStringPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return null != returnItem ? Convert.ToString(returnItem) : null;
        }


        #endregion

        #region ExecuteEnumPropertyGet<T>

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name) where T : struct, IConvertible
        {
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteEnumPropertyGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteEnumPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static T ExecuteEnumPropertyGetInternal<T>(this Core value, ICOMObject caller, string name, object[] validatedArgs) where T : struct, IConvertible
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            object intReturnItem = Convert.ToInt32(returnItem);
            T newObject = (T)intReturnItem;
            return newObject;
        }

        #endregion

        #region ExecuteStructPropertyGet<T>

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name) where T : struct
        {
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteStructPropertyGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteStructPropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static T ExecuteStructPropertyGetInternal<T>(this Core value, ICOMObject caller, string name, object[] validatedArgs) where T : struct
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            object intReturnItem = Convert.ToInt32(returnItem);
            T newObject = (T)intReturnItem;
            return newObject;
        }

        #endregion

        #region ExecuteReferencePropertyGet

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name) 
        {
            return ExecuteReferencePropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteReferencePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteReferencePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteReferencePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteReferencePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteReferencePropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static ICOMObject ExecuteReferencePropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);         
            ICOMObject newObject = value.CreateObjectFromComProxy(caller, returnItem, true);
            return newObject;
        }

        #endregion

        #region ExecuteReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static T ExecuteReferencePropertyGetInternal<T>(this Core value, ICOMObject caller, string name, object[] validatedArgs) where T : class, ICOMObject
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            T newObject = value.CreateObjectFromComProxy(caller, returnItem, true) as T;
            return newObject;
        }

        #endregion

        #region ExecuteBaseReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteBaseReferencePropertyGetInternal<T>(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static T ExecuteBaseReferencePropertyGetInternal<T>(this Core value, ICOMObject caller, string name, object[] validatedArgs) where T : class, ICOMObject
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            T newObject = value.CreateObjectFromComProxy(caller, returnItem, true) as T;
            return newObject;
        }

        #endregion

        #region ExecuteKnownReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType) where T : class, ICOMObject
        {
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteKnownReferencePropertyGetInternal<T>(value, caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static T ExecuteKnownReferencePropertyGetInternal<T>(this Core value, ICOMObject caller, string name, Type knownType, object[] validatedArgs) where T : class, ICOMObject
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            return value.CreateKnownObjectFromComProxy(caller, returnItem, knownType) as T;
        }

        #endregion

        #region ExecuteVariantPropertyGet

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteVariantPropertyGetInternal(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument);
            return ExecuteVariantPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2);
            return ExecuteVariantPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3);
            return ExecuteVariantPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4);
            return ExecuteVariantPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static object ExecuteVariantPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return ExecuteVariantPropertyGetInternal(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="validatedArgs">validated arguments as any</param>
        internal static object ExecuteVariantPropertyGetInternal(this Core value, ICOMObject caller, string name, object[] validatedArgs)
        {
            object returnItem = value.Invoker.PropertyGet(caller, name, validatedArgs);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = value.CreateObjectFromComProxy(caller, returnItem, false);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        #endregion
    }
}
