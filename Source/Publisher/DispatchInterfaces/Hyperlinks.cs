using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Hyperlinks 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class Hyperlinks : COMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Hyperlink>
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
                    _type = typeof(Hyperlinks);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Hyperlinks(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PublisherApi.Hyperlink this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Hyperlink>(this, "Item", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, index);
			}
		}

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
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		/// <param name="textToDisplay">optional string TextToDisplay = </param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage, object pageID, object textToDisplay)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, new object[]{ text, address, relativePage, pageID, textToDisplay });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, text);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, text, address);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, text, address, relativePage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="text">NetOffice.PublisherApi.TextRange text</param>
		/// <param name="address">optional string Address = </param>
		/// <param name="relativePage">optional NetOffice.PublisherApi.Enums.PbHlinkTargetType RelativePage = 1</param>
		/// <param name="pageID">optional Int32 PageID = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Hyperlink Add(NetOffice.PublisherApi.TextRange text, object address, object relativePage, object pageID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Hyperlink>(this, "Add", NetOffice.PublisherApi.Hyperlink.LateBindingApiWrapperType, text, address, relativePage, pageID);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Hyperlink>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Hyperlink>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Hyperlink>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Hyperlink>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public IEnumerator<NetOffice.PublisherApi.Hyperlink> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Hyperlink item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}