using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VBIDEApi
{
	///<summary>
	/// DispatchInterface _dispVBComponentsEvents 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _dispVBComponentsEvents : COMObject
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
                    _type = typeof(_dispVBComponentsEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _dispVBComponentsEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBComponentsEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemAdded(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent);
			Invoker.Method(this, "ItemAdded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemRemoved(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent);
			Invoker.Method(this, "ItemRemoved", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		/// <param name="oldName">string OldName</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemRenamed(NetOffice.VBIDEApi.VBComponent vBComponent, string oldName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent, oldName);
			Invoker.Method(this, "ItemRenamed", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemSelected(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent);
			Invoker.Method(this, "ItemSelected", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemActivated(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent);
			Invoker.Method(this, "ItemActivated", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent VBComponent</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemReloaded(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBComponent);
			Invoker.Method(this, "ItemReloaded", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}