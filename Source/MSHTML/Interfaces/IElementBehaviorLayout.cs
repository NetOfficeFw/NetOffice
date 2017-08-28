using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IElementBehaviorLayout 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IElementBehaviorLayout : COMObject
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
                    _type = typeof(IElementBehaviorLayout);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IElementBehaviorLayout(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IElementBehaviorLayout(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IElementBehaviorLayout(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="sizeContent">tagSIZE sizeContent</param>
		/// <param name="pptTranslateBy">tagPOINT pptTranslateBy</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		/// <param name="psizeProposed">tagSIZE psizeProposed</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetSize(Int32 dwFlags, tagSIZE sizeContent, tagPOINT pptTranslateBy, tagPOINT pptTopLeft, tagSIZE psizeProposed)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetSize", new object[]{ dwFlags, sizeContent, pptTranslateBy, pptTopLeft, psizeProposed });
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetLayoutInfo()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetLayoutInfo");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetPosition(Int32 lFlags, tagPOINT pptTopLeft)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetPosition", lFlags, pptTopLeft);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="psizeIn">tagSIZE psizeIn</param>
		/// <param name="prcOut">tagRECT prcOut</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 MapSize(tagSIZE psizeIn, out tagRECT prcOut)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			prcOut = new NetOffice.MSHTMLApi.tagRECT();
			object[] paramsArray = Invoker.ValidateParamsArray(psizeIn, prcOut);
			object returnItem = Invoker.MethodReturn(this, "MapSize", paramsArray, modifiers);
			prcOut = (tagRECT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
