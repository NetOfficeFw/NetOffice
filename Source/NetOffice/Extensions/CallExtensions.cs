using System;
using System.Linq;
using NetOffice;
using NetOffice.Exceptions;

namespace NetOffice.Extensions.Calling
{
    /// <summary>
    /// ICOMObject Call Extensions
    /// </summary>
    public static class CallExtensions
    {
        /// <summary>
        /// Invoke ICOMObject method by name.
        /// </summary>
        /// <remarks>Should be called when dealing with optional arguments results in ugly code because this method can handle so-called named arguments</remarks>
        /// <exception cref="NetOfficeCOMException">Unable to complete the call</exception>
        /// <param name="comObject">target ICOMObject instance</param>
        /// <param name="name">name of the method</param>
        /// <param name="arg1">optional argument 1</param>
        /// <param name="arg2">optional argument 2</param>
        /// <param name="arg3">optional argument 3</param>
        /// <param name="arg4">optional argument 4</param>
        /// <param name="arg5">optional argument 5</param>
        /// <param name="arg6">optional argument 6</param>
        /// <param name="arg7">optional argument 7</param>
        /// <param name="arg8">optional argument 8</param>
        /// <param name="arg9">optional argument 9</param>
        /// <param name="arg10">optional argument 10</param>
        /// <param name="arg11">optional argument 11</param>
        /// <param name="arg12">optional argument 12</param>
        /// <param name="arg13">optional argument 13</param>
        /// <param name="arg14">optional argument 14</param>
        /// <param name="arg15">optional argument 15</param>
        /// <param name="arg16">optional argument 16</param>
        /// <param name="arg17">optional argument 17</param>
        /// <param name="arg18">optional argument 18</param>
        /// <param name="arg19">optional argument 19</param>
        /// <param name="arg20">optional argument 20</param>
        /// <param name="arg21">optional argument 21</param>
        /// <param name="arg22">optional argument 22</param>
        /// <param name="arg23">optional argument 23</param>
        /// <param name="arg24">optional argument 24</param>
        /// <param name="arg25">optional argument 25</param>
        /// <param name="arg26">optional argument 26</param>
        /// <param name="arg27">optional argument 27</param>
        /// <param name="arg28">optional argument 28</param>
        /// <param name="arg29">optional argument 29</param>
        /// <param name="arg30">optional argument 30</param>
        /// <param name="arg31">optional argument 31</param>
        /// <param name="arg32">optional argument 32</param>
        public static void Method(this ICOMObject comObject, string name,
                                  object arg1 = null, object arg2 = null, object arg3 = null, object arg4 = null,
                                  object arg5 = null, object arg6 = null, object arg7 = null, object arg8 = null,
                                  object arg9 = null, object arg10 = null, object arg11 = null, object arg12 = null,
                                  object arg13 = null, object arg14 = null, object arg15 = null, object arg16 = null,
                                  object arg17 = null, object arg18 = null, object arg19 = null, object arg20 = null,
                                  object arg21 = null, object arg22 = null, object arg23 = null, object arg24 = null,
                                  object arg25 = null, object arg26 = null, object arg27 = null, object arg28 = null,
                                  object arg29 = null, object arg30 = null, object arg31 = null, object arg32 = null)
        {
            object[] args = new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8,
                                           arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16,
                                           arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24,
                                           arg25, arg26, arg27, arg28, arg29, arg30, arg31, arg32};

            object lastEmptyArgument = args.LastOrDefault(e => e == null || e == Type.Missing);

            int arrayLength = 0;
            object[] argArray = null;

            foreach (object item in args)
            {
                if (item == lastEmptyArgument)
                    break;
                arrayLength++;
            }

            argArray = new object[arrayLength];
            for (int i = 0; i < arrayLength; i++)
                argArray[i] = args[i];

            CoreMethodExtensions.ExecuteMethod(comObject.Factory, comObject, name, argArray);
        }

        /// <summary>
        /// Invoke ICOMObject method by name.
        /// </summary>
        /// <remarks>Should be called when dealing with optional arguments results in ugly code because this method can handle so-called named arguments</remarks>
        /// <exception cref="NetOfficeCOMException">Unable to complete the call</exception>
        /// <param name="comObject">target ICOMObject instance</param>
        /// <param name="name">name of the method</param>
        /// <param name="arg1">optional argument 1</param>
        /// <param name="arg2">optional argument 2</param>
        /// <param name="arg3">optional argument 3</param>
        /// <param name="arg4">optional argument 4</param>
        /// <param name="arg5">optional argument 5</param>
        /// <param name="arg6">optional argument 6</param>
        /// <param name="arg7">optional argument 7</param>
        /// <param name="arg8">optional argument 8</param>
        /// <param name="arg9">optional argument 9</param>
        /// <param name="arg10">optional argument 10</param>
        /// <param name="arg11">optional argument 11</param>
        /// <param name="arg12">optional argument 12</param>
        /// <param name="arg13">optional argument 13</param>
        /// <param name="arg14">optional argument 14</param>
        /// <param name="arg15">optional argument 15</param>
        /// <param name="arg16">optional argument 16</param>
        /// <param name="arg17">optional argument 17</param>
        /// <param name="arg18">optional argument 18</param>
        /// <param name="arg19">optional argument 19</param>
        /// <param name="arg20">optional argument 20</param>
        /// <param name="arg21">optional argument 21</param>
        /// <param name="arg22">optional argument 22</param>
        /// <param name="arg23">optional argument 23</param>
        /// <param name="arg24">optional argument 24</param>
        /// <param name="arg25">optional argument 25</param>
        /// <param name="arg26">optional argument 26</param>
        /// <param name="arg27">optional argument 27</param>
        /// <param name="arg28">optional argument 28</param>
        /// <param name="arg29">optional argument 29</param>
        /// <param name="arg30">optional argument 30</param>
        /// <param name="arg31">optional argument 31</param>
        /// <param name="arg32">optional argument 32</param>
        /// <returns>result of T</returns>
        public static T MethodGet<T>(this ICOMObject comObject, string name, 
                                    object arg1 = null, object arg2 = null, object arg3 = null, object arg4 = null,
                                    object arg5 = null, object arg6 = null, object arg7 = null, object arg8 = null,
                                    object arg9 = null, object arg10 = null, object arg11 = null, object arg12 = null,
                                    object arg13 = null, object arg14 = null, object arg15 = null, object arg16 = null,
                                    object arg17 = null, object arg18 = null, object arg19 = null, object arg20 = null,
                                    object arg21 = null, object arg22 = null, object arg23 = null, object arg24 = null,
                                    object arg25 = null, object arg26 = null, object arg27 = null, object arg28 = null,
                                    object arg29 = null, object arg30 = null, object arg31 = null, object arg32 = null)
        {
            object[] args = new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8,
                                           arg9, arg10, arg11, arg12, arg13, arg14, arg15, arg16,
                                           arg17, arg18, arg19, arg20, arg21, arg22, arg23, arg24,
                                           arg25, arg26, arg27, arg28, arg29, arg30, arg31, arg32};

            object lastEmptyArgument = args.LastOrDefault(e => e == null || e == Type.Missing);

            int arrayLength = 0;
            object[] argArray = null;

            foreach (object item in args)
            {
                if (item == lastEmptyArgument)
                    break;
                arrayLength++;
            }

            argArray = new object[arrayLength];
            for (int i = 0; i < arrayLength; i++)
                argArray[i] = args[i];

            object result = CoreMethodExtensions.ExecuteVariantMethodGet(comObject.Factory, comObject, name, argArray);
            if (result is T)
                return (T)result;
            else
                return default(T);
        }
    }
}