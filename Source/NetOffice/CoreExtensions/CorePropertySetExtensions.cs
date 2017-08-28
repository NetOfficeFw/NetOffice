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
    /// Provides top-off Core/Invoker set property services to shrink caller code in Api assemblies and give more refactoring possibilies
    /// </summary>
    public static class CorePropertySetExtensions
    {
        #region ExecutePropertySet

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }
        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }


        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecutePropertySet(this Core value, ICOMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);
            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion

        #region ExecuteValuePropertySet

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument,  newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }
        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }


        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteValuePropertySet(this Core value, ICOMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);
            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion

        #region ExecuteValuePropertySet<T>
     
        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument)
        {
            object[] arg = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, arg);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteValuePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);

            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion

        #region ExecuteEnumPropertySet

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a value property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteEnumPropertySet(this Core value, ICOMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);
            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion
        
        #region ExecuteReferencePropertySet

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteReferencePropertySet(this Core value, ICOMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);
            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion

        #region ExecuteReferencePropertySet<T>

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a reference property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteReferencePropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object[] paramsArray) where T:class,ICOMObject
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);

            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion
       
        #region ExecuteVariantPropertySet

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue)
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument)
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue, object argument1, object argument2, object argument3, object argument4)
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteVariantPropertySet(this Core value, ICOMObject caller, string name, object newValue, object[] paramsArray)
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);

            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion

        #region ExecuteVariantPropertySet<T>

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="argument1">argument as any</param>
        /// <param name="argument2">argument as any</param>
        /// <param name="argument3">argument as any</param>
        /// <param name="argument4">argument as any</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object argument1, object argument2, object argument3, object argument4) where T : class, ICOMObject
        {
            object[] args = Invoker.ValidateParamsArray(argument1, argument2, argument3, argument4, newValue);
            value.Invoker.PropertySet(caller, name, args);
        }

        /// <summary>
        /// Execute a variant property set
        /// </summary>
        /// <param name="value">core invoker</param>
        /// <param name="caller">calling instance</param>
        /// <param name="name">property name</param>
        /// <param name="newValue">value to set</param>
        /// <param name="paramsArray">arguments as any</param>
        public static void ExecuteVariantPropertySet<T>(this Core value, ICOMObject caller, string name, T newValue, object[] paramsArray) where T : class, ICOMObject
        {
            object[] newParamsArray = new object[paramsArray.Length + 1];
            for (int i = 0; i < paramsArray.Length; i++)
                newParamsArray[i] = paramsArray[i];
            newParamsArray[newParamsArray.Length - 1] = Invoker.ValidateParam(newValue);

            value.Invoker.PropertySet(caller, name, newParamsArray);
        }

        #endregion
    }
}
