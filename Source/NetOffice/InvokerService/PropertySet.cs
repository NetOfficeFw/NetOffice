using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.InvokerService
{
    /// <summary>
    /// Internal COMObject Property Set Invoker Service
    /// </summary>
    public static partial class InvokeInternal
    {
        #region ExecutePropertySet

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue)
        {
            object[] args = new object[] { newValue };
            ExecutePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecutePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecutePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecutePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4, newValue };
            ExecutePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecutePropertySet(this COMObject caller, string name, object newValue, object[] args)
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecutePropertySetInternal(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecutePropertySetInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length -1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if(!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteValuePropertySet

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue)
        {
            object[] args = new object[] { newValue };
            ExecuteValuePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecuteValuePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteValuePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteValuePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4, newValue };
            ExecuteValuePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteValuePropertySet(this COMObject caller, string name, object newValue, object[] args)
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteValuePropertySetInternal(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteValuePropertySetInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteValuePropertySet<T>

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue)
        {
            object[] args = new object[] { newValue };
            ExecuteValuePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecuteValuePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteValuePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteValuePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[]{ argument1, argument2, argument3, argument4, newValue};
            ExecuteValuePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteValuePropertySet<T>(this COMObject caller, string name, T newValue, object[] args)
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteValuePropertySetInternal<T>(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteValuePropertySetInternal<T>(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteEnumPropertySet

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue)
        {
            object[] args = new object[] { newValue };
            ExecuteEnumPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecuteEnumPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteEnumPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteEnumPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[]{ argument1, argument2, argument3, argument4, newValue};
            ExecuteEnumPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteEnumPropertySet(this COMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[null != paramsArray ? paramsArray.Length + 1 : 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteEnumPropertySetInternal(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteEnumPropertySetInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteReferencePropertySet

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue)
        {
            object[] args = new object[] { newValue };
            ExecuteReferencePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecuteReferencePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteReferencePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteReferencePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[]{ argument1, argument2, argument3, argument4, newValue};
            ExecuteReferencePropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteReferencePropertySet(this COMObject caller, string name, object newValue, object[] args)
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteReferencePropertySetInternal(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteReferencePropertySetInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteReferencePropertySet<T>

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue) where T : class, ICOMObject
        {
            object[] args = new object[] { newValue };
            ExecuteReferencePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue, object argument) where T : class, ICOMObject
        {
            object[] args = new object[] { argument, newValue };
            ExecuteReferencePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteReferencePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteReferencePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4, newValue };
            ExecuteReferencePropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteReferencePropertySet<T>(this COMObject caller, string name, T newValue, object[] args) where T : class, ICOMObject
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteReferencePropertySetInternal<T>(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteReferencePropertySetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteVariantPropertySet

            /// <summary>
            /// Execute a variant property set
            /// </summary>
            /// <param name="caller">calling instance</param>
            /// <param name="name">property name</param>
            /// <param name="newValue">value to set</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue)
        {
            object[] args = new object[] { newValue };
            ExecuteVariantPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue, object argument)
        {
            object[] args = new object[] { argument, newValue };
            ExecuteVariantPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteVariantPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteVariantPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4, newValue };
            ExecuteVariantPropertySetInternal(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteVariantPropertySet(this COMObject caller, string name, object newValue, object[] args)
        {
            object[] newParamsArray = new object[null != args ? args.Length + 1 : 1];
            for (int i = 0; i < args.Length; i++)
                newParamsArray[i] = args[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteVariantPropertySetInternal(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteVariantPropertySetInternal(this COMObject caller, string name, object[] args)
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion

        #region ExecuteVariantPropertySet<T>

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue) where T : class, ICOMObject
        {
            object[] args = new object[] { newValue };
            ExecuteVariantPropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue, object argument) where T : class, ICOMObject
        {
            object[] args = new object[] { argument, newValue };
            ExecuteVariantPropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, newValue };
            ExecuteVariantPropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, newValue };
            ExecuteVariantPropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = new object[] { argument1, argument2, argument3, argument4, newValue };
            ExecuteVariantPropertySetInternal<T>(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteVariantPropertySet<T>(this COMObject caller, string name, T newValue, object[] paramsArray) where T : class, ICOMObject
        {
            object[] newParamsArray = new object[null != paramsArray ? paramsArray.Length + 1 : 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = newValue;
            ExecuteVariantPropertySetInternal<T>(caller, name, newParamsArray);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="args">arguments as any</param>
        public static void ExecuteVariantPropertySetInternal<T>(this COMObject caller, string name, object[] args) where T : class, ICOMObject
        {
            try
            {
                caller.BeforeExecute(ExecuteMode.PropertySet, name, args);
                caller.ExecutePropertySet(name, args);
                caller.AfterExecute(ExecuteMode.PropertySet, name, args, null != args && args.Length > 0 ? args[args.Length - 1] : null);
            }
            catch (Exception exception)
            {
                bool continueAnyway = false;
                object continueResult = null;
                caller.ExecutionError(ExecuteMode.PropertySet, name, args, exception, ref continueAnyway, ref continueResult);
                if (!continueAnyway)
                    throw;
            }
        }

        #endregion
    }
}
