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
	/// DispatchInterface SharedWorkspaceMembers 
	/// SupportByVersion Office, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers"/> </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SharedWorkspaceMembers : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceMember>
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
                    _type = typeof(SharedWorkspaceMembers);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SharedWorkspaceMembers(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SharedWorkspaceMembers(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceMembers(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.SharedWorkspaceMember this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceMember>(this, "Item", NetOffice.OfficeApi.SharedWorkspaceMember.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers.Count"/> </remarks>
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
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers.Parent"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers.ItemCountExceeded"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool ItemCountExceeded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ItemCountExceeded");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers.Add"/> </remarks>
		/// <param name="email">string email</param>
		/// <param name="domainName">string domainName</param>
		/// <param name="displayName">string displayName</param>
		/// <param name="role">optional object role</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceMember Add(string email, string domainName, string displayName, object role)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceMember>(this, "Add", NetOffice.OfficeApi.SharedWorkspaceMember.LateBindingApiWrapperType, email, domainName, displayName, role);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceMembers.Add"/> </remarks>
		/// <param name="email">string email</param>
		/// <param name="domainName">string domainName</param>
		/// <param name="displayName">string displayName</param>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceMember Add(string email, string domainName, string displayName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceMember>(this, "Add", NetOffice.OfficeApi.SharedWorkspaceMember.LateBindingApiWrapperType, email, domainName, displayName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceMember>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceMember>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceMember>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceMember>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.SharedWorkspaceMember> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.SharedWorkspaceMember item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}