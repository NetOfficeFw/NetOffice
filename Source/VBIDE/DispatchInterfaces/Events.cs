using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface Events 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Events : COMObject
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
                    _type = typeof(Events);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Events(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Events(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Events(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.ReferencesEvents get_ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.ReferencesEvents>(this, "ReferencesEvents", NetOffice.VBIDEApi.ReferencesEvents.LateBindingApiWrapperType, vBProject);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ReferencesEvents
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject vBProject</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_ReferencesEvents")]
		public NetOffice.VBIDEApi.ReferencesEvents ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
		{
			return get_ReferencesEvents(vBProject);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="commandBarControl">object commandBarControl</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.CommandBarEvents get_CommandBarEvents(object commandBarControl)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CommandBarEvents>(this, "CommandBarEvents", NetOffice.VBIDEApi.CommandBarEvents.LateBindingApiWrapperType, commandBarControl);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_CommandBarEvents
		/// </summary>
		/// <param name="commandBarControl">object commandBarControl</param>
		[SupportByVersion("VBIDE", 12,14,5.3), Redirect("get_CommandBarEvents")]
		public NetOffice.VBIDEApi.CommandBarEvents CommandBarEvents(object commandBarControl)
		{
			return get_CommandBarEvents(commandBarControl);
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
