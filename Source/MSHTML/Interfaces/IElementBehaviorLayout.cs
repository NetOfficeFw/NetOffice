using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// Interface IElementBehaviorLayout 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IElementBehaviorLayout : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="dwFlags">Int32 dwFlags</param>
		/// <param name="sizeContent">tagSIZE sizeContent</param>
		/// <param name="pptTranslateBy">tagPOINT pptTranslateBy</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		/// <param name="psizeProposed">tagSIZE psizeProposed</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetSize(Int32 dwFlags, tagSIZE sizeContent, tagPOINT pptTranslateBy, tagPOINT pptTopLeft, tagSIZE psizeProposed)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwFlags, sizeContent, pptTranslateBy, pptTopLeft, psizeProposed);
			object returnItem = Invoker.MethodReturn(this, "GetSize", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetLayoutInfo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetLayoutInfo", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="lFlags">Int32 lFlags</param>
		/// <param name="pptTopLeft">tagPOINT pptTopLeft</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetPosition(Int32 lFlags, tagPOINT pptTopLeft)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lFlags, pptTopLeft);
			object returnItem = Invoker.MethodReturn(this, "GetPosition", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="psizeIn">tagSIZE psizeIn</param>
		/// <param name="prcOut">tagRECT prcOut</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MapSize(tagSIZE psizeIn, out tagRECT prcOut)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			prcOut = new NetOffice.MSHTMLApi.tagRECT();
			object[] paramsArray = Invoker.ValidateParamsArray(psizeIn, prcOut);
			object returnItem = Invoker.MethodReturn(this, "MapSize", paramsArray);
			prcOut = (tagRECT)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}