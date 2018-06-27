using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupContainer2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IMarkupContainer2 : IMarkupContainer, NetOffice.MSHTMLApi.IMarkupContainer2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupContainer2);
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
                    _type = typeof(IMarkupContainer2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupContainer2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="ppChangeLog">NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog</param>
		/// <param name="fForward">Int32 fForward</param>
		/// <param name="fBackward">Int32 fBackward</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 CreateChangeLog(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog, Int32 fForward, Int32 fBackward)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);
			ppChangeLog = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, ppChangeLog, fForward, fBackward);
			object returnItem = Invoker.MethodReturn(this, "CreateChangeLog", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppChangeLog = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLChangeLog>(this, paramsArray[1], typeof(NetOffice.MSHTMLApi.IHTMLChangeLog));
            else
                ppChangeLog = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="pdwCookie">Int32 pdwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RegisterForDirtyRange(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out Int32 pdwCookie)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pdwCookie = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, pdwCookie);
			object returnItem = Invoker.MethodReturn(this, "RegisterForDirtyRange", paramsArray, modifiers);
			pdwCookie = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 UnRegisterForDirtyRange(Int32 dwCookie)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "UnRegisterForDirtyRange", dwCookie);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		/// <param name="pIPointerBegin">NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin</param>
		/// <param name="pIPointerEnd">NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetAndClearDirtyRange(Int32 dwCookie, NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin, NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetAndClearDirtyRange", dwCookie, pIPointerBegin, pIPointerEnd);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetVersionNumber()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetVersionNumber");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElementMaster">NetOffice.MSHTMLApi.IHTMLElement ppElementMaster</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetMasterElement(out NetOffice.MSHTMLApi.IHTMLElement ppElementMaster)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppElementMaster = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppElementMaster);
			object returnItem = Invoker.MethodReturn(this, "GetMasterElement", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppElementMaster = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElementMaster = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		#endregion

		#pragma warning restore
	}
}

