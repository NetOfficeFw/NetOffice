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
	/// DispatchInterface Permission 
	/// SupportByVersion Office, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861518.aspx </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class Permission : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.UserPermission>
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
                    _type = typeof(Permission);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Permission(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Permission(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Permission(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.UserPermission this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.UserPermission>(this, "Item", NetOffice.OfficeApi.UserPermission.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865565.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862116.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool EnableTrustedBrowser
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableTrustedBrowser");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableTrustedBrowser", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861383.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865228.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862755.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string RequestPermissionURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RequestPermissionURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RequestPermissionURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860910.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string PolicyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PolicyName");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860601.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string PolicyDescription
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PolicyDescription");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864131.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool StoreLicenses
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StoreLicenses");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StoreLicenses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864905.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string DocumentAuthor
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DocumentAuthor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DocumentAuthor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863690.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool PermissionFromPolicy
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PermissionFromPolicy");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
		/// <param name="userId">string userId</param>
		/// <param name="permission">optional object permission</param>
		/// <param name="expirationDate">optional object expirationDate</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.UserPermission Add(string userId, object permission, object expirationDate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.UserPermission>(this, "Add", NetOffice.OfficeApi.UserPermission.LateBindingApiWrapperType, userId, permission, expirationDate);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
		/// <param name="userId">string userId</param>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.UserPermission Add(string userId)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.UserPermission>(this, "Add", NetOffice.OfficeApi.UserPermission.LateBindingApiWrapperType, userId);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863139.aspx </remarks>
		/// <param name="userId">string userId</param>
		/// <param name="permission">optional object permission</param>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.UserPermission Add(string userId, object permission)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.UserPermission>(this, "Add", NetOffice.OfficeApi.UserPermission.LateBindingApiWrapperType, userId, permission);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864678.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void ApplyPolicy(string fileName)
		{
			 Factory.ExecuteMethod(this, "ApplyPolicy", fileName);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861135.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void RemoveAll()
		{
			 Factory.ExecuteMethod(this, "RemoveAll");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.UserPermission>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.UserPermission>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.UserPermission>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.UserPermission>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.UserPermission> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.UserPermission item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}