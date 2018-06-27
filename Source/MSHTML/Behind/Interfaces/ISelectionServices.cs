using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface ISelectionServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISelectionServices : COMObject, NetOffice.MSHTMLApi.ISelectionServices
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
                    _contractType = typeof(NetOffice.MSHTMLApi.ISelectionServices);
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
                    _type = typeof(ISelectionServices);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ISelectionServices() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SetSelectionType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SetSelectionType", eType, pIListener);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppIContainer">NetOffice.MSHTMLApi.IMarkupContainer ppIContainer</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppIContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppIContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppIContainer);
			object returnItem = Invoker.MethodReturn(this, "GetMarkupContainer", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppIContainer = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IMarkupContainer>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.IMarkupContainer));
            else
                ppIContainer = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStart">NetOffice.MSHTMLApi.IMarkupPointer pIStart</param>
		/// <param name="pIEnd">NetOffice.MSHTMLApi.IMarkupPointer pIEnd</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.ISegment ppISegmentAdded</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AddSegment(NetOffice.MSHTMLApi.IMarkupPointer pIStart, NetOffice.MSHTMLApi.IMarkupPointer pIEnd, out NetOffice.MSHTMLApi.ISegment ppISegmentAdded)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppISegmentAdded = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIStart, pIEnd, ppISegmentAdded);
			object returnItem = Invoker.MethodReturn(this, "AddSegment", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppISegmentAdded = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ISegment>(this, paramsArray[2], typeof(NetOffice.MSHTMLApi.ISegment));
            else
                ppISegmentAdded = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AddElementSegment(NetOffice.MSHTMLApi.IHTMLElement pIElement, out NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppISegmentAdded = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIElement, ppISegmentAdded);
			object returnItem = Invoker.MethodReturn(this, "AddElementSegment", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppISegmentAdded = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.IElementSegment>(this, paramsArray[1], typeof(NetOffice.MSHTMLApi.IElementSegment));
            else
                ppISegmentAdded = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);         
        }

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.ISegment pISegment</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RemoveSegment(NetOffice.MSHTMLApi.ISegment pISegment)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveSegment", pISegment);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppISelectionServicesListener">NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetSelectionServicesListener(out NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppISelectionServicesListener = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppISelectionServicesListener);
			object returnItem = Invoker.MethodReturn(this, "GetSelectionServicesListener", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppISelectionServicesListener = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ISelectionServicesListener>(this, paramsArray[0], typeof(NetOffice.MSHTMLApi.ISelectionServicesListener));
            else
                ppISelectionServicesListener = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

