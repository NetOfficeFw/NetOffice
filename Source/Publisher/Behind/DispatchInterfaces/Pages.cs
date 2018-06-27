using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface Pages 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	public class Pages : COMObject, NetOffice.PublisherApi.Pages
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PublisherApi.Pages);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Pages() : base()
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
		public virtual NetOffice.PublisherApi.Page this[Int32 item]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Page>(this, "Item", typeof(NetOffice.PublisherApi.Page), item);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
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
		public virtual NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add10", typeof(NetOffice.PublisherApi.Page), count, after, duplicateObjectsOnPage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Page Add10(Int32 count, Int32 after)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add10", typeof(NetOffice.PublisherApi.Page), count, after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddWizardPage10(Int32 after, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddWizardPage10", after, pageType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddWizardPage10(Int32 after)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddWizardPage10", after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="pageID">Int32 pageID</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Page FindByPageID(Int32 pageID)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "FindByPageID", typeof(NetOffice.PublisherApi.Page), pageID);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage, object addHyperlinkToWebNavBar)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", typeof(NetOffice.PublisherApi.Page), count, after, duplicateObjectsOnPage, addHyperlinkToWebNavBar);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Page Add(Int32 count, Int32 after)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", typeof(NetOffice.PublisherApi.Page), count, after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="count">Int32 count</param>
		/// <param name="after">Int32 after</param>
		/// <param name="duplicateObjectsOnPage">optional Int32 DuplicateObjectsOnPage = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Page Add(Int32 count, Int32 after, object duplicateObjectsOnPage)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Page>(this, "Add", typeof(NetOffice.PublisherApi.Page), count, after, duplicateObjectsOnPage);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		/// <param name="addHyperlinkToWebNavBar">optional bool AddHyperlinkToWebNavBar = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddWizardPage(Int32 after, object pageType, object addHyperlinkToWebNavBar)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddWizardPage", after, pageType, addHyperlinkToWebNavBar);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddWizardPage(Int32 after)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddWizardPage", after);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="after">Int32 after</param>
		/// <param name="pageType">optional NetOffice.PublisherApi.Enums.PbWizardPageType PageType = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void AddWizardPage(Int32 after, object pageType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddWizardPage", after, pageType);
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
        public virtual IEnumerator<NetOffice.PublisherApi.Page> GetEnumerator()
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

