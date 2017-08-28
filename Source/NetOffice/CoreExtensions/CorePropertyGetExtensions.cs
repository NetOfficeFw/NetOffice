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
            return ExecuteObjectPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteObjectPropertyGet(value, caller, name, args);
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
            return ExecuteObjectPropertyGet(value, caller, name, args);
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
            return ExecuteObjectPropertyGet(value, caller, name, args);
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
            return ExecuteObjectPropertyGet(value, caller, name, args);
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
            return value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteInt16PropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteInt16PropertyGet(value, caller, name, args);
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
            return ExecuteInt16PropertyGet(value, caller, name, args);
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
            return ExecuteInt16PropertyGet(value, caller, name, args);
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
            return ExecuteInt16PropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteInt32PropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteInt32PropertyGet(value, caller, name, args);
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
            return ExecuteInt32PropertyGet(value, caller, name, args);
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
            return ExecuteInt32PropertyGet(value, caller, name, args);
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
            return ExecuteInt32PropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteInt64PropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteInt64PropertyGet(value, caller, name, args);
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
            return ExecuteInt64PropertyGet(value, caller, name, args);
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
            return ExecuteInt64PropertyGet(value, caller, name, args);
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
            return ExecuteInt64PropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteUIntPtrPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteUIntPtrPropertyGet(value, caller, name, args);
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
            return ExecuteUIntPtrPropertyGet(value, caller, name, args);
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
            return ExecuteUIntPtrPropertyGet(value, caller, name, args);
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
            return ExecuteUIntPtrPropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteFloatPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteFloatPropertyGet(value, caller, name, args);
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
            return ExecuteFloatPropertyGet(value, caller, name, args);
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
            return ExecuteFloatPropertyGet(value, caller, name, args);
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
            return ExecuteFloatPropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteDoublePropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteDoublePropertyGet(value, caller, name, args);
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
            return ExecuteDoublePropertyGet(value, caller, name, args);
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
            return ExecuteDoublePropertyGet(value, caller, name, args);
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
            return ExecuteDoublePropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteSinglePropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteSinglePropertyGet(value, caller, name, args);
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
            return ExecuteSinglePropertyGet(value, caller, name, args);
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
            return ExecuteSinglePropertyGet(value, caller, name, args);
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
            return ExecuteSinglePropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteDateTimePropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteDateTimePropertyGet(value, caller, name, args);
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
            return ExecuteDateTimePropertyGet(value, caller, name, args);
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
            return ExecuteDateTimePropertyGet(value, caller, name, args);
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
            return ExecuteDateTimePropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteBoolPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteBoolPropertyGet(value, caller, name, args);
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
            return ExecuteBoolPropertyGet(value, caller, name, args);
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
            return ExecuteBoolPropertyGet(value, caller, name, args);
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
            return ExecuteBoolPropertyGet(value, caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static bool ExecuteBoolPropertyGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteStringPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteStringPropertyGet(value, caller, name, args);
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
            return ExecuteStringPropertyGet(value, caller, name, args);
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
            return ExecuteStringPropertyGet(value, caller, name, args);
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
            return ExecuteStringPropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteEnumPropertyGet<T>(value, caller, name, _emptyParams);
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
            return ExecuteEnumPropertyGet<T>(value, caller, name, args);
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
            return ExecuteEnumPropertyGet<T>(value, caller, name, args);
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
            return ExecuteEnumPropertyGet<T>(value, caller, name, args);
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
            return ExecuteEnumPropertyGet<T>(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteStructPropertyGet<T>(value, caller, name, _emptyParams);
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
            return ExecuteStructPropertyGet<T>(value, caller, name, args);
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
            return ExecuteStructPropertyGet<T>(value, caller, name, args);
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
            return ExecuteStructPropertyGet<T>(value, caller, name, args);
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
            return ExecuteStructPropertyGet<T>(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
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
            return ExecuteReferencePropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteReferencePropertyGet(value, caller, name, args);
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
            return ExecuteReferencePropertyGet(value, caller, name, args);
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
            return ExecuteReferencePropertyGet(value, caller, name, args);
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
            return ExecuteReferencePropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
            ICOMObject newObject = value.CreateObjectFromComProxy(caller, returnItem);
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
            return ExecuteReferencePropertyGet<T>(value, caller, name, _emptyParams);
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
            return ExecuteReferencePropertyGet<T>(value, caller, name, args);
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
            return ExecuteReferencePropertyGet<T>(value, caller, name, args);
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
            return ExecuteReferencePropertyGet<T>(value, caller, name, args);
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
            return ExecuteReferencePropertyGet<T>(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
            T newObject = value.CreateObjectFromComProxy(caller, returnItem) as T;
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
            return ExecuteKnownReferencePropertyGet<T>(value, caller, name, knownType, _emptyParams);
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
            return ExecuteKnownReferencePropertyGet<T>(value, caller, name, knownType, args);
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
            return ExecuteKnownReferencePropertyGet<T>(value, caller, name, knownType, args);
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
            return ExecuteKnownReferencePropertyGet<T>(value, caller, name, knownType, args);
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
            return ExecuteKnownReferencePropertyGet<T>(value, caller, name, knownType, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
            T newObject = value.CreateKnownObjectFromComProxy(caller, returnItem, knownType) as T;
            return newObject;
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
            return ExecuteVariantPropertyGet(value, caller, name, _emptyParams);
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
            return ExecuteVariantPropertyGet(value, caller, name, args);
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
            return ExecuteVariantPropertyGet(value, caller, name, args);
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
            return ExecuteVariantPropertyGet(value, caller, name, args);
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
            return ExecuteVariantPropertyGet(value, caller, name, args);
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
            object returnItem = value.Invoker.PropertyGet(caller, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = value.CreateObjectFromComProxy(caller, returnItem);
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
