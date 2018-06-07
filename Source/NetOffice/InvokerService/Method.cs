using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NetOffice.InvokerService
{
    /// <summary>
    /// Internal COMObject Method Invoker Service
    /// </summary>
    public static partial class InvokeInternal
    {
        #region Fields

        private static object[] _emptyParams = new object[0];

        #endregion

        #region ExecuteMethod

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static void ExecuteMethod(this COMObject caller, string name)
        {
            ExecuteMethodInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteMethod( this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3, argument4, argument5 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
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
        public static void ExecuteMethod(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            ExecuteMethodInternal(caller, name, new object[] { argument1, argument2, argument3,
                argument4, argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteMethod(this COMObject caller, string name, object[] args)
        {
            ExecuteMethodInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static void ExecuteMethodInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.Method, name, args);
                caller.CallMethod(name, args);
                caller.AfterExecute(ExecuteMode.Method, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.Method, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        /// <summary>
        /// Execute a method without return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        ///  <param name="modifiers">optional modifiers to deal with ref and out arguments</param>
        public static void ExecuteMethodExtended(this COMObject caller, string name, object[] args, ParameterModifier[] modifiers)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.Method, name, args);
                caller.CallMethod(name, args, modifiers);
                caller.AfterExecute(ExecuteMode.Method, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.Method, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteObjectMethodGet

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name)
        {
            return ExecuteObjectMethodGet(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argumen1 as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
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
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteObjectMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteObjectMethodGet(this COMObject caller, string name, object[] args)
        {
            return caller.CallMethodGet(name, args);
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static object ExecuteObjectMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(object);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return result;
        }

        #endregion

        #region ExecuteInt16MethodGet

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name)
        {
            return ExecuteInt16MethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
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
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteInt16MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static Int16 ExecuteInt16MethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteInt16MethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static Int16 ExecuteInt16MethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Int16);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToInt16(result) : default(Int16);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int16)result;
        }

        #endregion

        #region ExecuteInt32MethodGet

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name)
        {
            return ExecuteInt32MethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
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
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteInt32MethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static Int32 ExecuteInt32MethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteInt32MethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static Int32 ExecuteInt32MethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Int32);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToInt32(result) : default(Int32);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int32)result;
        }

        /// <summary>
        /// Execute a method with Int32 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        /// <param name="modifiers">optional modifiers to deal with ref and out arguments</param>
        public static Int32 ExecuteInt32MethodGetExtended(this COMObject caller, string name, object[] args, ParameterModifier[] modifiers)
        {
            object result = default(Int32);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args, modifiers);
                result = null != result ? Convert.ToInt32(result) : default(Int32);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int32)result;
        }

        #endregion

        #region ExecuteDoubleMethodGet

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name)
        {
            return ExecuteDoubleMethodGet(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
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
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteDoubleMethodGet(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static double ExecuteDoubleMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteDoubleMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static double ExecuteDoubleMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(double);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToDouble(result) : default(double);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (double)result;
        }

        #endregion

        #region ExecuteSingleMethodGet

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name)
        {
            return ExecuteSingleMethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
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
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteSingleMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static Single ExecuteSingleMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteSingleMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static Single ExecuteSingleMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Single);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToSingle(result) : default(Single);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Single)result;
        }

        #endregion

        #region ExecuteBooleanMethodGet

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name)
        {
            return ExecuteBoolMethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
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
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteBoolMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static bool ExecuteBoolMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteBoolMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static bool ExecuteBoolMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(bool);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToBoolean(result) : default(bool);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (bool)result;
        }

        #endregion

        #region ExecuteDateTimeMethodGet

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
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
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static DateTime ExecuteDateTimeMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteDateTimeMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static DateTime ExecuteDateTimeMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(DateTime);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToDateTime(result) : default(DateTime);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (DateTime)result;
        }

        #endregion

        #region ExecuteStringMethodGet

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name)
        {
            return ExecuteStringMethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with string bool value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
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
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3,
            object argument4, object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteStringMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static string ExecuteStringMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteStringMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        internal static string ExecuteStringMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(string);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                result = null != result ? Convert.ToString(result) : default(string);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (string)result;
        }

        #endregion

        #region ExecuteEnumMethodGet

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary> 
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name, object argument) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteEnumMethodGet<T>(this COMObject caller, string name, object[] args) where T : struct, IConvertible
        {
            return ExecuteEnumMethodGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a method get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        internal static T ExecuteEnumMethodGetInternal<T>(this COMObject caller, string name, object[] args) where T : struct, IConvertible
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                object intReturnItem = Convert.ToInt32(result);
                result = (T)intReturnItem;
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        /// <summary>
        /// Execute a method get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        ///  <param name="modifiers">optional modifiers to deal with optional arguments</param>
        public static T ExecuteEnumMethodGetExtended<T>(this COMObject caller, string name, object[] args, ParameterModifier[] modifiers) where T : struct, IConvertible
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args, modifiers);
                object intReturnItem = Convert.ToInt32(result);
                result = (T)intReturnItem;
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteStructMethodGet

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name, object argument) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteStructMethodGet<T>(this COMObject caller, string name, object[] args) where T : struct
        {
            return ExecuteStructMethodGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a method get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        internal static T ExecuteStructMethodGetInternal<T>(this COMObject caller, string name, object[] args) where T : struct
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteReferenceMethodGet

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
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
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteReferenceMethodGet<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            return ExecuteReferenceMethodGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a method get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        internal static T ExecuteReferenceMethodGetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, true);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteBaseReferenceMethodGet

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGet<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
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
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteBaseReferenceMethodGet<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            return ExecuteBaseReferenceMethodGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        internal static T ExecuteBaseReferenceMethodGetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, false);
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteKnownReferenceMethodGet

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType = null) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, _emptyParams);
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
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
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
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
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteKnownReferenceMethodGet<T>(this COMObject caller, string name, Type knownType, object[] args) where T : class, ICOMObject
        {
            return ExecuteKnownReferenceMethodGetInternal<T>(caller, name, knownType, args);
        }
        
        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="args">arguments as any</param>
        internal static T ExecuteKnownReferenceMethodGetInternal<T>(this COMObject caller, string name, Type knownType, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateKnownObjectFromComProxy(caller, result, knownType) as T;
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }


        /// <summary>
        /// Execute a method with known reference return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="knownType">type of T to increase performance</param>
        /// <param name="args">arguments as any</param>
        /// <param name="modifiers">modifiers to deal with ref and out arguments</param>
        public static T ExecuteKnownReferenceMethodGetExtended<T>(this COMObject caller, string name, Type knownType, object[] args, ParameterModifier[] modifiers) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args, modifiers);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateKnownObjectFromComProxy(caller, result, knownType) as T;
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteVariantMethodGet

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name)
        {
            return ExecuteVariantMethodGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        /// <param name="argument5">argument as any</param>
        /// <param name="argument6">argument as any</param>
        /// <param name="argument7">argument as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
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
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4,
            object argument5, object argument6, object argument7, object argument8)
        {
            return ExecuteVariantMethodGetInternal(caller, name, new object[] { argument1, argument2, argument3, argument4,
                argument5, argument6, argument7, argument8 });
        }

        /// <summary>
        /// Execute a method with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">method name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteVariantMethodGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteVariantMethodGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a method get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        internal static object ExecuteVariantMethodGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(object);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args);
                if ((null != result) && (result is MarshalByRefObject))
                {
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, false);
                }
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return result;
        }


        /// <summary>
        /// Execute a method get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        /// <param name="modifiers">optional modifiers to deal with ref and out arguments</param>
        internal static object ExecuteVariantMethodGetExtended(this COMObject caller, string name, object[] args, ParameterModifier[] modifiers)
        {
            object result = default(object);
            try
            {
                caller.BeforeExecute(ExecuteMode.MethodGet, name, args);
                result = caller.CallMethodGet(name, args, modifiers);
                if ((null != result) && (result is MarshalByRefObject))
                {
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, false);
                }
                caller.AfterExecute(ExecuteMode.MethodGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.MethodGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return result;
        }
        #endregion
    }
}
