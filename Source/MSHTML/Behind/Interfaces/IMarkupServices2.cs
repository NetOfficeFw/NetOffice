using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IMarkupServices2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IMarkupServices2 : IMarkupServices, NetOffice.MSHTMLApi.IMarkupServices2
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IMarkupServices2);
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
                    _type = typeof(IMarkupServices2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMarkupServices2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="hglobalHTML">_userHGLOBAL hglobalHTML</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="pContext">NetOffice.MSHTMLApi.IMarkupContainer pContext</param>
		/// <param name="ppContainerResult">NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult</param>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ParseGlobalEx(_userHGLOBAL hglobalHTML, Int32 dwFlags, NetOffice.MSHTMLApi.IMarkupContainer pContext, out NetOffice.MSHTMLApi.IMarkupContainer ppContainerResult, NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,false,false);
			ppContainerResult = null;
			object[] paramsArray = Invoker.ValidateParamsArray(hglobalHTML, dwFlags, pContext, ppContainerResult, pPointerStart, pPointerFinish);
			object returnItem = Invoker.MethodReturn(this, "ParseGlobalEx", paramsArray, modifiers);
            if (paramsArray[3] is MarshalByRefObject)
                ppContainerResult = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupContainer>(this, paramsArray[3], typeof(NetOffice.MSHTMLApi.IMarkupContainer));
            else
                ppContainerResult = null;
            return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pPointerStart">NetOffice.MSHTMLApi.IMarkupPointer pPointerStart</param>
		/// <param name="pPointerFinish">NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish</param>
		/// <param name="pPointerTarget">NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget</param>
		/// <param name="pPointerStatus">NetOffice.MSHTMLApi.IMarkupPointer pPointerStatus</param>
		/// <param name="ppElemFailBottom">NetOffice.MSHTMLApi.IHTMLElement ppElemFailBottom</param>
		/// <param name="ppElemFailTop">NetOffice.MSHTMLApi.IHTMLElement ppElemFailTop</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 ValidateElements(NetOffice.MSHTMLApi.IMarkupPointer pPointerStart, NetOffice.MSHTMLApi.IMarkupPointer pPointerFinish, NetOffice.MSHTMLApi.IMarkupPointer pPointerTarget, NetOffice.MSHTMLApi.IMarkupPointer pPointerStatus, out NetOffice.MSHTMLApi.IHTMLElement ppElemFailBottom, out NetOffice.MSHTMLApi.IHTMLElement ppElemFailTop)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,true);
			ppElemFailBottom = null;
			ppElemFailTop = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointerStart, pPointerFinish, pPointerTarget, pPointerStatus, ppElemFailBottom, ppElemFailTop);
			object returnItem = Invoker.MethodReturn(this, "ValidateElements", paramsArray, modifiers);
            if (paramsArray[4] is MarshalByRefObject)
                ppElemFailBottom = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[4], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElemFailBottom = null;
            if (paramsArray[5] is MarshalByRefObject)
                ppElemFailTop = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHTMLElement>(this, paramsArray[5], typeof(NetOffice.MSHTMLApi.IHTMLElement));
            else
                ppElemFailTop = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSegmentList">NetOffice.MSHTMLApi.ISegmentList pSegmentList</param>
		/// <param name="dwFlags">Int32 dwFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SaveSegmentsToClipboard(NetOffice.MSHTMLApi.ISegmentList pSegmentList, Int32 dwFlags)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SaveSegmentsToClipboard", pSegmentList, dwFlags);
		}

		#endregion

		#pragma warning restore
	}
}

