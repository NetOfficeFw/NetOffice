using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IElementSegment 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IElementSegment : ISegment, NetOffice.MSHTMLApi.IElementSegment
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IElementSegment);
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
                    _type = typeof(IElementSegment);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IElementSegment() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIElement">NetOffice.MSHTMLApi.IHTMLElement ppIElement</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetElement(out NetOffice.MSHTMLApi.IHTMLElement ppIElement)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppIElement = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppIElement);
			object returnItem = Invoker.MethodReturn(this, "GetElement", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppIElement = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppIElement = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="fPrimary">Int32 fPrimary</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetPrimary(Int32 fPrimary)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetPrimary", fPrimary);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pfPrimary">Int32 pfPrimary</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 IsPrimary(out Int32 pfPrimary)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pfPrimary = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pfPrimary);
			object returnItem = Invoker.MethodReturn(this, "IsPrimary", paramsArray, modifiers);
			pfPrimary = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

