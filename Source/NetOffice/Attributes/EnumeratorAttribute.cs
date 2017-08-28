using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetOffice.Attributes
{    
    /// <summary>
     /// Indicates how an enumerator want invoke internally
     /// </summary>
    public enum EnumeratorInvoke
    {
        /// <summary>
        /// Invoke as property
        /// </summary>
        Property = 0,

        /// <summary>
        /// Invoke as method
        /// </summary>
        Method = 1,
       
        /// <summary>
        /// Underlying instance doesnt have an enumerator. Enumerator is spend by NetOffice instead by using this+Count.
        /// Using a custom enumerator may cause performance/memory dropdown when heavy amount of data is returned
        /// or fetching result cause remote calls - for example get thousands of outlook mail items 
        /// from an Exchange Server through NetOffice custom enumerator isn't a good idea. 
        /// </summary>
        Custom = 2
    }

    /// <summary>
    /// Enumerator Result Kind
    /// </summary>
    public enum Enumerator
    {
        /// <summary>
        /// Unknown/may mixed
        /// </summary>
        Variant = 0,

        /// <summary>
        /// Returns reference types
        /// </summary>
        Reference = 1,

        /// <summary>
        /// Returns scalar/value types like bool or int(Boolean or Integer in Visual Basic) and also System.String
        /// </summary>
        Value = 2
    }

    /// <summary>
    /// Enumerator Implementation Details
    /// </summary>
    public class EnumeratorAttribute : System.Attribute
    {
        /// <summary>
        /// Return Kind
        /// </summary>
        public readonly Enumerator Result;

        /// <summary>
        /// Internal Invoke Kind
        /// </summary>
        public readonly EnumeratorInvoke Invoke;

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="result">return kind</param>
        /// <param name="invoke">internal invoke call</param>
        public EnumeratorAttribute(Enumerator result, EnumeratorInvoke invoke)
        {
            Result = result;
            Invoke = invoke;
        }
    }
}
