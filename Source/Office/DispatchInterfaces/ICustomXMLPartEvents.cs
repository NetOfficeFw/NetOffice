using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface ICustomXMLPartEvents 
	/// SupportByVersion Office, 12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ICustomXMLPartEvents : COMObject
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
                    _type = typeof(ICustomXMLPartEvents);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ICustomXMLPartEvents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ICustomXMLPartEvents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode NewNode</param>
		/// <param name="inUndoRedo">bool InUndoRedo</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void NodeAfterInsert(NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(newNode, inUndoRedo);
			Invoker.Method(this, "NodeAfterInsert", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="oldParentNode">NetOffice.OfficeApi.CustomXMLNode OldParentNode</param>
		/// <param name="oldNextSibling">NetOffice.OfficeApi.CustomXMLNode OldNextSibling</param>
		/// <param name="inUndoRedo">bool InUndoRedo</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void NodeAfterDelete(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode oldParentNode, NetOffice.OfficeApi.CustomXMLNode oldNextSibling, bool inUndoRedo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, oldParentNode, oldNextSibling, inUndoRedo);
			Invoker.Method(this, "NodeAfterDelete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="oldNode">NetOffice.OfficeApi.CustomXMLNode OldNode</param>
		/// <param name="newNode">NetOffice.OfficeApi.CustomXMLNode NewNode</param>
		/// <param name="inUndoRedo">bool InUndoRedo</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void NodeAfterReplace(NetOffice.OfficeApi.CustomXMLNode oldNode, NetOffice.OfficeApi.CustomXMLNode newNode, bool inUndoRedo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldNode, newNode, inUndoRedo);
			Invoker.Method(this, "NodeAfterReplace", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}