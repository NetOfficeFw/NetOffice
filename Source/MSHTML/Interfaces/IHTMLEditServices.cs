using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLEditServices 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface), BaseType]
 	public class IHTMLEditServices : COMObject
	{
		#pragma warning disable

		#region Type Information

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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLEditServices(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLEditServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(string progId) : base(progId)
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
		public Int32 AddDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			return Factory.ExecuteInt32MethodGet(this, "AddDesigner", pIDesigner);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 RemoveDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			return Factory.ExecuteInt32MethodGet(this, "RemoveDesigner", pIDesigner);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIContainer">NetOffice.MSHTMLApi.IMarkupContainer pIContainer</param>
		/// <param name="ppSelSvc">NetOffice.MSHTMLApi.ISelectionServices ppSelSvc</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetSelectionServices(NetOffice.MSHTMLApi.IMarkupContainer pIContainer, out NetOffice.MSHTMLApi.ISelectionServices ppSelSvc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppSelSvc = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIContainer, ppSelSvc);
			object returnItem = Invoker.MethodReturn(this, "GetSelectionServices", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppSelSvc = new NetOffice.MSHTMLApi.ISelectionServices(this, paramsArray[1]);
            else
                ppSelSvc = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 MoveToSelectionAnchor(NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveToSelectionAnchor", pIStartAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 MoveToSelectionEnd(NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveToSelectionEnd", pIEndAnchor);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pStart">NetOffice.MSHTMLApi.IMarkupPointer pStart</param>
		/// <param name="pEnd">NetOffice.MSHTMLApi.IMarkupPointer pEnd</param>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 SelectRange(NetOffice.MSHTMLApi.IMarkupPointer pStart, NetOffice.MSHTMLApi.IMarkupPointer pEnd, NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType)
		{
			return Factory.ExecuteInt32MethodGet(this, "SelectRange", pStart, pEnd, eType);
		}

		#endregion

		#pragma warning restore
	}
}
