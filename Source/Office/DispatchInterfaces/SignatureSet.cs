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
	/// DispatchInterface SignatureSet 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet"/> </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SignatureSet : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.Signature>
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
                    _type = typeof(SignatureSet);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SignatureSet(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SignatureSet(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SignatureSet(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.Count"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="iSig">Int32 iSig</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.Signature this[Int32 iSig]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Signature>(this, "Item", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType, iSig);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.Parent"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.CanAddSignatureLine"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool CanAddSignatureLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanAddSignatureLine");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.Subset"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoSignatureSubset Subset
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSignatureSubset>(this, "Subset");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Subset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.ShowSignaturesPane"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool ShowSignaturesPane
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSignaturesPane");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSignaturesPane", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Signature Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "Add", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Commit()
		{
			 Factory.ExecuteMethod(this, "Commit");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.AddNonVisibleSignature"/> </remarks>
		/// <param name="varSigProv">optional object varSigProv</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Signature AddNonVisibleSignature(object varSigProv)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddNonVisibleSignature", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType, varSigProv);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.AddNonVisibleSignature"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Signature AddNonVisibleSignature()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddNonVisibleSignature", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.AddSignatureLine"/> </remarks>
		/// <param name="varSigProv">optional object varSigProv</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Signature AddSignatureLine(object varSigProv)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddSignatureLine", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType, varSigProv);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SignatureSet.AddSignatureLine"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Signature AddSignatureLine()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Signature>(this, "AddSignatureLine", NetOffice.OfficeApi.Signature.LateBindingApiWrapperType);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.Signature>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.Signature>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.Signature>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Signature>

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.Signature> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.Signature item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}