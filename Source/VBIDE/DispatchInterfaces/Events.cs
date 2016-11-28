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
	/// DispatchInterface Events 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Events : COMObject
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
                    _type = typeof(Events);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.ReferencesEvents get_ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(vBProject);
			object returnItem = Invoker.PropertyGet(this, "ReferencesEvents", paramsArray);
			NetOffice.VBIDEApi.ReferencesEvents newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.ReferencesEvents.LateBindingApiWrapperType) as NetOffice.VBIDEApi.ReferencesEvents;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_ReferencesEvents
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.ReferencesEvents ReferencesEvents(NetOffice.VBIDEApi.VBProject vBProject)
		{
			return get_ReferencesEvents(vBProject);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		/// <param name="commandBarControl">object CommandBarControl</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VBIDEApi.CommandBarEvents get_CommandBarEvents(object commandBarControl)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(commandBarControl);
			object returnItem = Invoker.PropertyGet(this, "CommandBarEvents", paramsArray);
			NetOffice.VBIDEApi.CommandBarEvents newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.VBIDEApi.CommandBarEvents.LateBindingApiWrapperType) as NetOffice.VBIDEApi.CommandBarEvents;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Alias for get_CommandBarEvents
		/// </summary>
		/// <param name="commandBarControl">object CommandBarControl</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
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