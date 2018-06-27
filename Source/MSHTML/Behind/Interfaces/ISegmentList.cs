using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface ISegmentList 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISegmentList : COMObject, NetOffice.MSHTMLApi.ISegmentList
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSHTMLApi.ISegmentList);
                return _contractType;
            }
        }
        private static Type _contractType;


		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ISegmentList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISegmentList() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIIter">NetOffice.MSHTMLApi.ISegmentListIterator ppIIter</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateIterator(out NetOffice.MSHTMLApi.ISegmentListIterator ppIIter)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppIIter = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppIIter);
			object returnItem = Invoker.MethodReturn(this, "CreateIterator", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppIIter = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ISegmentListIterator>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.ISegmentListIterator));
            else
                ppIIter = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="peType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetType(out NetOffice.MSHTMLApi.Enums._SELECTION_TYPE peType)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peType = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peType);
			object returnItem = Invoker.MethodReturn(this, "GetType", paramsArray, modifiers);
			peType = (NetOffice.MSHTMLApi.Enums._SELECTION_TYPE)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfEmpty">Int32 pfEmpty</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsEmpty(out Int32 pfEmpty)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfEmpty = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfEmpty);
			object returnItem = Invoker.MethodReturn(this, "IsEmpty", paramsArray, modifiers);
			pfEmpty = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

