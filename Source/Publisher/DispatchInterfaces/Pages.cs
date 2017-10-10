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
	/// DispatchInterface Pages 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class Pages : COMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Page>
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
                    _type = typeof(Pages);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Pages(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Pages(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="item">Int32 item</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PublisherApi.Page this[Int32 item]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Page>(this, "Item", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, item);
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
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add10", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, count, after, duplicateObjectsOnPage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add10", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, count, after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddWizardPage10(Int32 after, object pageType)
		{
			 Factory.ExecuteMethod(this, "AddWizardPage10", after, pageType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddWizardPage10(Int32 after)
		{
			 Factory.ExecuteMethod(this, "AddWizardPage10", after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pageID">Int32 pageID</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page FindByPageID(Int32 pageID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "FindByPageID", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, pageID);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage, object addHyperlinkToWebNavBar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, count, after, duplicateObjectsOnPage, addHyperlinkToWebNavBar);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, count, after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", NetOffice.PublisherApi.Page.LateBindingApiWrapperType, count, after, duplicateObjectsOnPage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after, object pageType, object addHyperlinkToWebNavBar)
		{
			 Factory.ExecuteMethod(this, "AddWizardPage", after, pageType, addHyperlinkToWebNavBar);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after)
		{
			 Factory.ExecuteMethod(this, "AddWizardPage", after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddWizardPage(Int32 after, object pageType)
		{
			 Factory.ExecuteMethod(this, "AddWizardPage", after, pageType);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Page>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Page>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Page>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Page>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public IEnumerator<NetOffice.PublisherApi.Page> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Page item in innerEnumerator)
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