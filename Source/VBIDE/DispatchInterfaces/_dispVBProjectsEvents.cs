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
	/// DispatchInterface _dispVBProjectsEvents 
	/// SupportByVersion VBIDE, 12,14,5.3
	///</summary>
	[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _dispVBProjectsEvents : COMObject
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
                    _type = typeof(_dispVBProjectsEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _dispVBProjectsEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _dispVBProjectsEvents(string progId) : base(progId)
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
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemAdded(NetOffice.VBIDEApi.VBProject vBProject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBProject);
			Invoker.Method(this, "ItemAdded", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemRemoved(NetOffice.VBIDEApi.VBProject vBProject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBProject);
			Invoker.Method(this, "ItemRemoved", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		/// <param name="oldName">string OldName</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemRenamed(NetOffice.VBIDEApi.VBProject vBProject, string oldName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBProject, oldName);
			Invoker.Method(this, "ItemRenamed", paramsArray);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// 
		/// </summary>
		/// <param name="vBProject">NetOffice.VBIDEApi.VBProject VBProject</param>
		[SupportByVersionAttribute("VBIDE", 12,14,5.3)]
		public void ItemActivated(NetOffice.VBIDEApi.VBProject vBProject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(vBProject);
			Invoker.Method(this, "ItemActivated", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}