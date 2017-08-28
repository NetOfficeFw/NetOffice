using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPainterEventInfo 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLPainterEventInfo : COMObject
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
                    _type = typeof(IHTMLPainterEventInfo);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLPainterEventInfo(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLPainterEventInfo(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainterEventInfo(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="plEventInfoFlags">Int32 plEventInfoFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetEventInfoFlags(out Int32 plEventInfoFlags)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			plEventInfoFlags = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(plEventInfoFlags);
			object returnItem = Invoker.MethodReturn(this, "GetEventInfoFlags", paramsArray, modifiers);
			plEventInfoFlags = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElement">NetOffice.MSHTMLApi.IHTMLElement ppElement</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetEventTarget(NetOffice.MSHTMLApi.IHTMLElement ppElement)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetEventTarget", ppElement);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 SetCursor(Int32 lPartID)
		{
			return Factory.ExecuteInt32MethodGet(this, "SetCursor", lPartID);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lPartID">Int32 lPartID</param>
		/// <param name="pbstrPart">string pbstrPart</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 StringFromPartID(Int32 lPartID, out string pbstrPart)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrPart = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(lPartID, pbstrPart);
			object returnItem = Invoker.MethodReturn(this, "StringFromPartID", paramsArray, modifiers);
			pbstrPart = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
