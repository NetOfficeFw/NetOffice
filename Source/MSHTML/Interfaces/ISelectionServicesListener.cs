using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface ISelectionServicesListener 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class ISelectionServicesListener : COMObject
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
                    _type = typeof(ISelectionServicesListener);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISelectionServicesListener(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISelectionServicesListener(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 BeginSelectionUndo()
		{
			return Factory.ExecuteInt32MethodGet(this, "BeginSelectionUndo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 EndSelectionUndo()
		{
			return Factory.ExecuteInt32MethodGet(this, "EndSelectionUndo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pIElementStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementStart</param>
		/// <param name="pIElementEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd</param>
		/// <param name="pIElementContentStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart</param>
		/// <param name="pIElementContentEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 OnSelectedElementExit(NetOffice.MSHTMLApi.IMarkupPointer pIElementStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd)
		{
			return Factory.ExecuteInt32MethodGet(this, "OnSelectedElementExit", pIElementStart, pIElementEnd, pIElementContentStart, pIElementContentEnd);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 OnChangeType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener)
		{
			return Factory.ExecuteInt32MethodGet(this, "OnChangeType", eType, pIListener);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pTypeDetail">string pTypeDetail</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetTypeDetail(out string pTypeDetail)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pTypeDetail = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pTypeDetail);
			object returnItem = Invoker.MethodReturn(this, "GetTypeDetail", paramsArray, modifiers);
			pTypeDetail = paramsArray[0] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
