using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface XMLSchemaReferences 
	/// SupportByVersion Word, 11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196203.aspx </remarks>
	public class XMLSchemaReferences : COMObject, NetOffice.WordApi.XMLSchemaReferences
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
                    _contractType = typeof(NetOffice.WordApi.XMLSchemaReferences);
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
                    _type = typeof(XMLSchemaReferences);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public XMLSchemaReferences() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838278.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192772.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821946.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835461.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool AutomaticValidation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutomaticValidation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutomaticValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool AllowSaveAsXMLWithoutValidation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AllowSaveAsXMLWithoutValidation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllowSaveAsXMLWithoutValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835525.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool HideValidationErrors
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HideValidationErrors");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HideValidationErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195673.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool IgnoreMixedContent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IgnoreMixedContent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IgnoreMixedContent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194590.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual bool ShowPlaceholderText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowPlaceholderText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowPlaceholderText", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.WordApi.XMLSchemaReference this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Item", typeof(NetOffice.WordApi.XMLSchemaReference), index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197504.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual void Validate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Validate");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840980.aspx </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName, object installForAllUsers)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", typeof(NetOffice.WordApi.XMLSchemaReference), namespaceURI, alias, fileName, installForAllUsers);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840980.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReference Add()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", typeof(NetOffice.WordApi.XMLSchemaReference));
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840980.aspx </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", typeof(NetOffice.WordApi.XMLSchemaReference), namespaceURI);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840980.aspx </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", typeof(NetOffice.WordApi.XMLSchemaReference), namespaceURI, alias);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840980.aspx </remarks>
		/// <param name="namespaceURI">optional object namespaceURI</param>
		/// <param name="alias">optional object alias</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public virtual NetOffice.WordApi.XMLSchemaReference Add(object namespaceURI, object alias, object fileName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.XMLSchemaReference>(this, "Add", typeof(NetOffice.WordApi.XMLSchemaReference), namespaceURI, alias, fileName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.XMLSchemaReference>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.XMLSchemaReference>

        /// <summary>
        /// SupportByVersion Word, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.WordApi.XMLSchemaReference> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.XMLSchemaReference item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

