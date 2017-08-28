using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IHTMLPainter 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IHTMLPainter : COMObject
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
                    _type = typeof(IHTMLPainter);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IHTMLPainter(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLPainter(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLPainter(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="rcBounds">tagRECT rcBounds</param>
		/// <param name="rcUpdate">tagRECT rcUpdate</param>
		/// <param name="lDrawFlags">Int32 lDrawFlags</param>
		/// <param name="hdc">_RemotableHandle hdc</param>
		/// <param name="pvDrawObject">object pvDrawObject</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 Draw(tagRECT rcBounds, tagRECT rcUpdate, Int32 lDrawFlags, _RemotableHandle hdc, object pvDrawObject)
		{
			return Factory.ExecuteInt32MethodGet(this, "Draw", new object[]{ rcBounds, rcUpdate, lDrawFlags, hdc, pvDrawObject });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="size">tagSIZE size</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 onresize(tagSIZE size)
		{
			return Factory.ExecuteInt32MethodGet(this, "onresize", size);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pInfo">_HTML_PAINTER_INFO pInfo</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetPainterInfo(out _HTML_PAINTER_INFO pInfo)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pInfo = new NetOffice.MSHTMLApi._HTML_PAINTER_INFO();
			object[] paramsArray = Invoker.ValidateParamsArray(pInfo);
			object returnItem = Invoker.MethodReturn(this, "GetPainterInfo", paramsArray, modifiers);
			pInfo = (_HTML_PAINTER_INFO)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pt">tagPOINT pt</param>
		/// <param name="pbHit">Int32 pbHit</param>
		/// <param name="plPartID">Int32 plPartID</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 HitTestPoint(tagPOINT pt, out Int32 pbHit, out Int32 plPartID)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			pbHit = 0;
			plPartID = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pt, pbHit, plPartID);
			object returnItem = Invoker.MethodReturn(this, "HitTestPoint", paramsArray, modifiers);
			pbHit = (Int32)paramsArray[1];
			plPartID = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
