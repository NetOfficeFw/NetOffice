using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface WebNavigationBarHyperlinks 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class WebNavigationBarHyperlinks : COMObject , IEnumerable<NetOffice.PublisherApi.Hyperlink>
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
                    _type = typeof(WebNavigationBarHyperlinks);
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public WebNavigationBarHyperlinks(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public WebNavigationBarHyperlinks(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PublisherApi.Hyperlink this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Item", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 0</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		/// <param name="textToDisplay">optional string TextToDisplay = </param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(object address, object relativePage, object pageID, object textToDisplay, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, new object[]{ address, relativePage, pageID, textToDisplay, index });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="address">optional string Address = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, address);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(object address, object relativePage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, address, relativePage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 0</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(object address, object relativePage, object pageID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, address, relativePage, pageID);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 0</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		/// <param name="textToDisplay">optional string TextToDisplay = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(object address, object relativePage, object pageID, object textToDisplay)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, address, relativePage, pageID, textToDisplay);
		}

		#endregion

       #region IEnumerable<NetOffice.PublisherApi.Hyperlink> Member
        
        /// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
       public IEnumerator<NetOffice.PublisherApi.Hyperlink> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.PublisherApi.Hyperlink item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersion Publisher, 14,15,16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}



