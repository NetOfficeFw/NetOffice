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
        
        In order to shrink the size of API assemblies as best as possible - we give 8 fixed argument overloads too.
        (API assemblies in 1.7.4.1 call fixed arguments overloads whenever its possible)
    */

    /// <summary>
    /// Provides top-off Core/Invoker method services to shrink caller code in Api assemblies and give more refactoring possibilies
    /// </summary> 
    public static class CoreMethodExtensions
    {
        #region Fields

        private static object[] _emptyParams = new object[0];

        #endregion

        #region ExecuteMethod

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name)
        {
            ExecuteMethod(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument)
        {
            ExecuteMethod(value, caller, name, new object[] { argument});
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, 
            object argument4)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3, argument4, argument5 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            ExecuteMethod(value, caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteMethod(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            value.Invoker.Method(caller, name, args);
        }

        #endregion
      
        #region ExecuteObjectMethodGet
     
        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteObjectMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5, object argument6)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteObjectMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static object ExecuteObjectMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            return value.Invoker.MethodReturn(caller, name, args);
        }

        #endregion

        #region ExecuteInt16MethodGet

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteInt16MethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteInt16MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Int16 ExecuteInt16MethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToInt16(returnItem) : (short)0;
        }

        #endregion
        
        #region ExecuteInt32MethodGet

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteInt32MethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }
        
        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteInt32MethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }
        
        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Int32 ExecuteInt32MethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToInt32(returnItem) : 0;
        }

        #endregion

        #region ExecuteDoubleMethodGet

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteDoubleMethodGet(value, caller, name, _emptyParams);
        }
       
        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, 
            object argument4, object argument5)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteDoubleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }
        
        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static double ExecuteDoubleMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToDouble(returnItem) : 0;
        }

        #endregion

        #region ExecuteSingleMethodGet
      
        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteSingleMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteSingleMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }
        
        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static Single ExecuteSingleMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToSingle(returnItem) : 0;
        }

        #endregion

        #region ExecuteBooleanMethodGet

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteBoolMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteBoolMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static bool ExecuteBoolMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToBoolean(returnItem) : false;
        }

        #endregion
        
        #region ExecuteDateTimeMethodGet

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteDateTimeMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToDateTime(returnItem) : default(DateTime);
        }

        #endregion
        
        #region ExecuteStringMethodGet

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteStringMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with string bool value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteStringMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static string ExecuteStringMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return null != returnItem ? Convert.ToString(returnItem) : null;
        }

        #endregion

        #region ExecuteEnumMethodGet
        
        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGet<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGet<T>(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGet<T>(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGet<T>(value, caller, name, new object[] { argument1, argument2 , argument3 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteEnumMethodGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : struct, IConvertible
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            T newObject = (T)returnItem;
            return newObject;
        }

        #endregion

        #region ExecuteStructMethodGet

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name) where T : struct
        {
            return ExecuteStructMethodGet<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : struct
        {
            return ExecuteStructMethodGet<T>(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : struct
        {
            return ExecuteStructMethodGet<T>(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : struct
        {
            return ExecuteStructMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct
        {
            return ExecuteStructMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteStructMethodGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : struct
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            T newObject = (T)returnItem;
            return newObject;
        }

        #endregion
        
        #region ExecuteReferenceMethodGet

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3});
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            T newObject = value.CreateObjectFromComProxy(caller, returnItem, true) as T;
            return newObject;
        }

        #endregion

        #region ExecuteBaseReferenceMethodGet

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            T newObject = value.CreateObjectFromComProxy(caller, returnItem, true) as T;
            return newObject;
        }

        #endregion
        
        #region ExecuteKnownReferenceMethodGet

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, _emptyParams);
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4, 
            object argument5) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGet<T>(value, caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="paramsArray">arguments as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this Core value, ICOMObject caller, string name, Type knownType, object[] paramsArray) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            return value.CreateKnownObjectFromComProxy(caller, returnItem, knownType) as T;
        }

        #endregion

        #region ExecuteVariantMethodGet
       
        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name)
        {
            return ExecuteVariantMethodGet(value, caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument });
        }
        
        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4, 
            object argument5)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        /// <param name="argument8">argument as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteVariantMethodGet(value, caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="paramsArray">arguments as any</param>
        public static object ExecuteVariantMethodGet(this Core value, ICOMObject caller, string name, object[] paramsArray)
        {
            object[] args = Invoker.ValidateParamsArray(paramsArray);
            object returnItem = value.Invoker.MethodReturn(caller, name, args);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = value.CreateObjectFromComProxy(caller, returnItem, true);
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