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
	/// DispatchInterface SharedWorkspaceFolders 
	/// SupportByVersion Office, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders"/> </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class SharedWorkspaceFolders : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceFolder>
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
                    _type = typeof(SharedWorkspaceFolders);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SharedWorkspaceFolders(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SharedWorkspaceFolders(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspaceFolders(string progId) : base(progId)
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
		public NetOffice.OfficeApi.SharedWorkspaceFolder this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceFolder>(this, "Item", NetOffice.OfficeApi.SharedWorkspaceFolder.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders.Count"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders.ItemCountExceeded"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders.Add"/> </remarks>
		/// <param name="folderName">string folderName</param>
		/// <param name="parentFolder">optional object parentFolder</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceFolder Add(string folderName, object parentFolder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceFolder>(this, "Add", NetOffice.OfficeApi.SharedWorkspaceFolder.LateBindingApiWrapperType, folderName, parentFolder);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspaceFolders.Add"/> </remarks>
		/// <param name="folderName">string folderName</param>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceFolder Add(string folderName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SharedWorkspaceFolder>(this, "Add", NetOffice.OfficeApi.SharedWorkspaceFolder.LateBindingApiWrapperType, folderName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceFolder>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceFolder>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.SharedWorkspaceFolder>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.SharedWorkspaceFolder>

        /// <summary>
        /// SupportByVersion Office, 11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.SharedWorkspaceFolder> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.SharedWorkspaceFolder item in innerEnumerator)
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