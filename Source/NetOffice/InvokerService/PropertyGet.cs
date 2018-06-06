using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.InvokerService
{
    /// <summary>
    /// Internal COMObject Property Get Invoker Service
    /// </summary>
    public static partial class InvokeInternal
    {
        #region ExecuteObjectPropertyGet

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name)
        {
            return ExecuteObjectPropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteObjectPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteObjectPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteObjectPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4};
            return ExecuteObjectPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteObjectPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteObjectPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with object return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteObjectPropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = null;
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);                
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return result;
        }

        #endregion

        #region ExecuteInt16PropertyGet

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name)
        {
            return ExecuteInt16PropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteInt16PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteInt16PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteInt16PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteInt16PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Int16 ExecuteInt16PropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteInt16PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int16 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Int16 ExecuteInt16PropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Int16);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToInt16(result) : default(Int16);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int16)result;
        }

        #endregion

        #region ExecuteInt32PropertyGet

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Int32 ExecuteInt32PropertyGet(this COMObject caller, string name)
        {
            return ExecuteInt32PropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Int32 ExecuteInt32PropertyGet( this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteInt32PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int32 ExecuteInt32PropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteInt32PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int32 ExecuteInt32PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteInt32PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int32 ExecuteInt32PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[]{ argument1, argument2, argument3, argument4};
            return ExecuteInt32PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Int32 ExecuteInt32PropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteInt32PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Int32 ExecuteInt32PropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Int32);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToInt32(result) : (Int32)0;
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int32)result;
        }

        #endregion

        #region ExecuteInt64PropertyGet

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name)
        {
            return ExecuteInt64PropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteInt64PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteInt64PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteInt64PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteInt64PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Int64 ExecuteInt64PropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteInt64PropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with Int64 return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">validated arguments</param>
        public static Int64 ExecuteInt64PropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Int64);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToInt64(result) : (Int64)0;
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Int64)result;
        }

        #endregion

        #region ExecuteUIntPtrPropertyGet

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name)
        {
            return ExecuteUIntPtrPropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteUIntPtrPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteUIntPtrPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteUIntPtrPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with UIntPtr return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteUIntPtrPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteUIntPtrPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with int return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static UIntPtr ExecuteUIntPtrPropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(UIntPtr);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? (UIntPtr)result : UIntPtr.Zero;
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (UIntPtr)result;
        }

        #endregion

        #region ExecuteFloatPropertyGet

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static float ExecuteFloatPropertyGet(this COMObject caller, string name)
        {
            return ExecuteFloatPropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static float ExecuteFloatPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteFloatPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static float ExecuteFloatPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteFloatPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static float ExecuteFloatPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteFloatPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static float ExecuteFloatPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteFloatPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Single ExecuteFloatPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteFloatPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with float return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Single ExecuteFloatPropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Single);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToSingle(result) : default(Single);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Single)result;
        }

        #endregion

        #region ExecuteDoublePropertyGet

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name)
        {
            return ExecuteDoublePropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteDoublePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteDoublePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteDoublePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteDoublePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static double ExecuteDoublePropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteDoublePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with double return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static double ExecuteDoublePropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(double);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToDouble(result) : default(double);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (double)result;
        }

        #endregion

        #region ExecuteSinglePropertyGet

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name)
        {
            return ExecuteSinglePropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteSinglePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteSinglePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteSinglePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteSinglePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Single ExecuteSinglePropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteSinglePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with single return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static Single ExecuteSinglePropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(Single);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToSingle(result) : default(Single);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (Single)result;
        }

        #endregion

        #region ExecuteDateTimePropertyGet

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name)
        {
            return ExecuteDateTimePropertyGet(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteDateTimePropertyGet(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteDateTimePropertyGet(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteDateTimePropertyGet(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteDateTimePropertyGet(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static DateTime ExecuteDateTimePropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteDateTimePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with DateTime return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static DateTime ExecuteDateTimePropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(DateTime);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToDateTime(result) : default(DateTime);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (DateTime)result;
        }

        #endregion

        #region ExecuteBytePropertyGet

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name)
        {
            return ExecuteBytePropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteBytePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteBytePropertyGetInternal( caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteBytePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteBytePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static byte ExecuteBytePropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteBytePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with byte return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static byte ExecuteBytePropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(byte);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToByte(result) : default(byte);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (byte)result;
        }

        #endregion

        #region ExecuteBoolPropertyGet

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name)
        {
            return ExecuteBoolPropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteBoolPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2};
            return ExecuteBoolPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteBoolPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteBoolPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static bool ExecuteBoolPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteBoolPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with bool return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static bool ExecuteBoolPropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(bool);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToBoolean(result) : default(bool);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (bool)result;
        }

        #endregion

        #region ExecuteStringPropertyGet

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name)
        {
            return ExecuteStringPropertyInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteStringPropertyInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2};
            return ExecuteStringPropertyInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteStringPropertyInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteStringPropertyInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static string ExecuteStringPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteStringPropertyInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with string return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static string ExecuteStringPropertyInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(string);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                result = null != result ? Convert.ToString(result) : default(string);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (string)result;
        }

        #endregion

        #region ExecuteEnumPropertyGet<T>

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name) where T : struct, IConvertible
        {
            return ExecuteEnumPropertyGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name, object argument) where T : struct, IConvertible
        {
            object[] args = new object[] { argument };
            return ExecuteEnumPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : struct, IConvertible
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteEnumPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : struct, IConvertible
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteEnumPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct, IConvertible
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteEnumPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteEnumPropertyGet<T>(this COMObject caller, string name, object[] args) where T : struct, IConvertible
        {
            return ExecuteEnumPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with enum return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteEnumPropertyGetInternal<T>(this COMObject caller, string name, object[] args) where T : struct, IConvertible
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                object intReturnItem = Convert.ToInt32(result);
                result = (T)intReturnItem;
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteStructPropertyGet<T>

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name) where T : struct
        {
            return ExecuteStructPropertyGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name, object argument) where T : struct
        {
            object[] args = new object[] { argument };
            return ExecuteStructPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : struct
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteStructPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : struct
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteStructPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : struct
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteStructPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteStructPropertyGet<T>(this COMObject caller, string name, object[] args) where T : struct
        {
            return ExecuteStructPropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with struct return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteStructPropertyGetInternal<T>(this COMObject caller, string name, object[] args) where T : struct
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteReferencePropertyGet

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name)
        {
            return ExecuteReferencePropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteReferencePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteReferencePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteReferencePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteReferencePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static ICOMObject ExecuteReferencePropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteReferencePropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static ICOMObject ExecuteReferencePropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(ICOMObject);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                if(!(result is ICOMObject))
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, true);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (ICOMObject)result;
        }

        #endregion

        #region ExecuteReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteReferencePropertyGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name, object argument) where T : class, ICOMObject
        {
            object[] args = new object[] { argument };
            return ExecuteReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteReferencePropertyGet<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            return ExecuteReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteReferencePropertyGetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, true);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteBaseReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name) where T : class, ICOMObject
        {
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name, object argument) where T : class, ICOMObject
        {
            object[] args = new object[] { argument };
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2 };
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteBaseReferencePropertyGet<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            return ExecuteBaseReferencePropertyGetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteBaseReferencePropertyGetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, false);
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteKnownReferencePropertyGet<T>

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this COMObject caller, string name, Type knownType = null) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferencePropertyGetInternal<T>(caller, name, knownType, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this COMObject caller, string name, Type knownType, object argument) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            object[] args = new object[] { argument };
            return ExecuteKnownReferencePropertyGetInternal<T>(caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            object[] args = new object[] { argument1, argument2 };
            return ExecuteKnownReferencePropertyGetInternal<T>(caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteKnownReferencePropertyGetInternal<T>(caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>( this COMObject caller, string name, Type knownType, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteKnownReferencePropertyGetInternal<T> (caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteKnownReferencePropertyGet<T>(this COMObject caller, string name, Type knownType, object[] args) where T : class, ICOMObject
        {
            if (null == knownType)
                knownType = typeof(T);
            return ExecuteKnownReferencePropertyGetInternal<T>(caller, name, knownType, args);
        }

        /// <summary>
        /// Execute a property get with known COM reference type return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="knownType">type of T - given to increase performance</param>
        /// <param name="args">arguments as any</param>
        public static T ExecuteKnownReferencePropertyGetInternal<T>(this COMObject caller, string name, Type knownType, object[] args) where T : class, ICOMObject
        {
            object result = default(T);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                if (!(result is ICOMObject))
                    result = caller.Factory.CreateKnownObjectFromComProxy(caller, result, knownType) as T;
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
                if (continueAnyway)
                    result = continueResult;
                else
                    throw;
            }
            return (T)result;
        }

        #endregion

        #region ExecuteVariantPropertyGet

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name)
        {
            return ExecuteVariantPropertyGetInternal(caller, name, _emptyParams);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument">argument as any</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name, object argument)
        {
            object[] args = new object[] { argument };
            return ExecuteVariantPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2};
            return ExecuteVariantPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3 };
            return ExecuteVariantPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4 };
            return ExecuteVariantPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteVariantPropertyGet(this COMObject caller, string name, object[] args)
        {
            return ExecuteVariantPropertyGetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property get with unknown return value
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static object ExecuteVariantPropertyGetInternal(this COMObject caller, string name, object[] args)
        {
            object result = default(object);
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertyGet, name, args);
                result = caller.ExecutePropertyGet(name, args);
                if ((null != result) && (result is MarshalByRefObject))
                {
                    result = caller.Factory.CreateObjectFromComProxy(caller, result, false);
                }
                caller.AfterExecute(ExecuteMode.PropertyGet, name, args, result);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertyGet, name, args, exception, ref continueAnyway, ref continueResult);
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
