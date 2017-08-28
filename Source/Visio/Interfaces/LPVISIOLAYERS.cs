using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOLAYERS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class LPVISIOLAYERS : COMObject , IEnumerable<NetOffice.VisioApi.IVLayer>
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
                    _type = typeof(LPVISIOLAYERS);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOLAYERS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOLAYERS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application", NetOffice.VisioApi.IVApplication.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document", NetOffice.VisioApi.IVDocument.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPage Page
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "Page", NetOffice.VisioApi.IVPage.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMaster Master
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "Master", NetOffice.VisioApi.IVMaster.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVLayer this[object nameOrIndex]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVLayer>(this, "Item", NetOffice.VisioApi.IVLayer.LateBindingApiWrapperType, nameOrIndex);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList", NetOffice.VisioApi.IVEventList.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVLayer get_ItemU(object nameOrIndex)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVLayer>(this, "ItemU", NetOffice.VisioApi.IVLayer.LateBindingApiWrapperType, nameOrIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemU
		/// </summary>
		/// <param name="nameOrIndex">object nameOrIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemU")]
		public NetOffice.VisioApi.IVLayer ItemU(object nameOrIndex)
		{
			return get_ItemU(nameOrIndex);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="layerName">string layerName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVLayer Add(string layerName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VisioApi.IVLayer>(this, "Add", NetOffice.VisioApi.IVLayer.LateBindingApiWrapperType, layerName);
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVLayer> Member
        
        /// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVLayer> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVLayer item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}



