using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOSTATUSBARS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class LPVISIOSTATUSBARS : COMObject , IEnumerable<NetOffice.VisioApi.IVStatusBar>
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
                    _type = typeof(LPVISIOSTATUSBARS);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOSTATUSBARS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSTATUSBARS(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVStatusBar this[Int32 lIndex]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVStatusBar>(this, "Item", NetOffice.VisioApi.IVStatusBar.LateBindingApiWrapperType, lIndex);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="lID">Int32 lID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVStatusBar get_ItemAtID(Int32 lID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVStatusBar>(this, "ItemAtID", NetOffice.VisioApi.IVStatusBar.LateBindingApiWrapperType, lID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ItemAtID
		/// </summary>
		/// <param name="lID">Int32 lID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ItemAtID")]
		public NetOffice.VisioApi.IVStatusBar ItemAtID(Int32 lID)
		{
			return get_ItemAtID(lID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVUIObject Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVUIObject>(this, "Parent", NetOffice.VisioApi.IVUIObject.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVStatusBar Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VisioApi.IVStatusBar>(this, "Add", NetOffice.VisioApi.IVStatusBar.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="lID">Int32 lID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVStatusBar AddAtID(Int32 lID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VisioApi.IVStatusBar>(this, "AddAtID", NetOffice.VisioApi.IVStatusBar.LateBindingApiWrapperType, lID);
		}

		#endregion

       #region IEnumerable<NetOffice.VisioApi.IVStatusBar> Member
        
        /// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
       public IEnumerator<NetOffice.VisioApi.IVStatusBar> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.VisioApi.IVStatusBar item in innerEnumerator)
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



