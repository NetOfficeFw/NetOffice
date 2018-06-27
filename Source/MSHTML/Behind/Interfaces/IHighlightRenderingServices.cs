using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHighlightRenderingServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHighlightRenderingServices : COMObject, NetOffice.MSHTMLApi.IHighlightRenderingServices
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHighlightRenderingServices);
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
                    _type = typeof(IHighlightRenderingServices);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHighlightRenderingServices() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		/// <param name="pIRenderStyle">NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle</param>
		/// <param name="ppISegment">NetOffice.MSHTMLApi.IHighlightSegment ppISegment</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AddSegment(NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd, NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle, out NetOffice.MSHTMLApi.IHighlightSegment ppISegment)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			ppISegment = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointerStart, pDispPointerEnd, pIRenderStyle, ppISegment);
			object returnItem = Invoker.MethodReturn(this, "AddSegment", paramsArray, modifiers);
            if (paramsArray[3] is MarshalByRefObject)
                ppISegment = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IHighlightSegment>(this, paramsArray[3], typeof(NetOffice.MSHTMLApi.IHighlightSegment));
            else
                ppISegment = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveSegmentToPointers(NetOffice.MSHTMLApi.IHighlightSegment pISegment, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveSegmentToPointers", pISegment, pDispPointerStart, pDispPointerEnd);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RemoveSegment(NetOffice.MSHTMLApi.IHighlightSegment pISegment)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveSegment", pISegment);
		}

		#endregion

		#pragma warning restore
	}
}

