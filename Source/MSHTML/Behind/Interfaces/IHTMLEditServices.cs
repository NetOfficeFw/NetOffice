using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IHTMLEditServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IHTMLEditServices : COMObject, NetOffice.MSHTMLApi.IHTMLEditServices
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IHTMLEditServices);
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
                    _type = typeof(IHTMLEditServices);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IHTMLEditServices() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 AddDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "AddDesigner", pIDesigner);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 RemoveDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RemoveDesigner", pIDesigner);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIContainer">NetOffice.MSHTMLApi.IMarkupContainer pIContainer</param>
		/// <param name="ppSelSvc">NetOffice.MSHTMLApi.ISelectionServices ppSelSvc</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetSelectionServices(NetOffice.MSHTMLApi.IMarkupContainer pIContainer, out NetOffice.MSHTMLApi.ISelectionServices ppSelSvc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppSelSvc = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIContainer, ppSelSvc);
			object returnItem = Invoker.MethodReturn(this, "GetSelectionServices", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppSelSvc = Factory.CreateKnownObjectFromComProxy<NetOffice.MSHTMLApi.ISelectionServices>(this, paramsArray[1], typeof(NetOffice.MSHTMLApi.ISelectionServices));
            else
                ppSelSvc = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToSelectionAnchor(NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToSelectionAnchor", pIStartAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 MoveToSelectionEnd(NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveToSelectionEnd", pIEndAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pStart">NetOffice.MSHTMLApi.IMarkupPointer pStart</param>
		/// <param name="pEnd">NetOffice.MSHTMLApi.IMarkupPointer pEnd</param>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 SelectRange(NetOffice.MSHTMLApi.IMarkupPointer pStart, NetOffice.MSHTMLApi.IMarkupPointer pEnd, NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "SelectRange", pStart, pEnd, eType);
		}

		#endregion

		#pragma warning restore
	}
}

