using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface _dispVBComponentsEvents 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _dispVBComponentsEvents : COMObject
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
                    _type = typeof(_dispVBComponentsEvents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _dispVBComponentsEvents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemAdded(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "ItemAdded", vBComponent);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemRemoved(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "ItemRemoved", vBComponent);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		/// <param name="oldName">string oldName</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemRenamed(NetOffice.VBIDEApi.VBComponent vBComponent, string oldName)
		{
			 Factory.ExecuteMethod(this, "ItemRenamed", vBComponent, oldName);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemSelected(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "ItemSelected", vBComponent);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemActivated(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "ItemActivated", vBComponent);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void ItemReloaded(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "ItemReloaded", vBComponent);
		}

		#endregion

		#pragma warning restore
	}
}
