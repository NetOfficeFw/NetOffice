using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface PickerResults 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class PickerResults : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PickerResult>
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
                    _type = typeof(PickerResults);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PickerResults(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PickerResults(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerResults(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.PickerResult this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerResult>(this, "Item", NetOffice.OfficeApi.PickerResult.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults.Count"/> </remarks>
		[SupportByVersion("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults.Add"/> </remarks>
		/// <param name="id">string id</param>
		/// <param name="displayName">string displayName</param>
		/// <param name="type">string type</param>
		/// <param name="sIPId">optional string SIPId = </param>
		/// <param name="itemData">optional object itemData</param>
		/// <param name="subItems">optional object subItems</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData, object subItems)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", NetOffice.OfficeApi.PickerResult.LateBindingApiWrapperType, new object[]{ id, displayName, type, sIPId, itemData, subItems });
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults.Add"/> </remarks>
		/// <param name="id">string id</param>
		/// <param name="displayName">string displayName</param>
		/// <param name="type">string type</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", NetOffice.OfficeApi.PickerResult.LateBindingApiWrapperType, id, displayName, type);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults.Add"/> </remarks>
		/// <param name="id">string id</param>
		/// <param name="displayName">string displayName</param>
		/// <param name="type">string type</param>
		/// <param name="sIPId">optional string SIPId = </param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", NetOffice.OfficeApi.PickerResult.LateBindingApiWrapperType, id, displayName, type, sIPId);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PickerResults.Add"/> </remarks>
		/// <param name="id">string id</param>
		/// <param name="displayName">string displayName</param>
		/// <param name="type">string type</param>
		/// <param name="sIPId">optional string SIPId = </param>
		/// <param name="itemData">optional object itemData</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResult Add(string id, string displayName, string type, object sIPId, object itemData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResult>(this, "Add", NetOffice.OfficeApi.PickerResult.LateBindingApiWrapperType, new object[]{ id, displayName, type, sIPId, itemData });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.PickerResult>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.PickerResult>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.PickerResult>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PickerResult>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.PickerResult> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.PickerResult item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}