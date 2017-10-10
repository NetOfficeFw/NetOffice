using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOVALIDATIONISSUES 
	/// SupportByVersion Visio, 14,15,16
	/// </summary>
	[SupportByVersion("Visio", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class LPVISIOVALIDATIONISSUES : COMObject, IEnumerableProvider<NetOffice.VisioApi.IVValidationIssue>
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
                    _type = typeof(LPVISIOVALIDATIONISSUES);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOVALIDATIONISSUES(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOVALIDATIONISSUES(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOVALIDATIONISSUES(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VisioApi.IVValidationIssue this[Int32 index]
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVValidationIssue>(this, "Item", index);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="issueID">Int32 issueID</param>
		[SupportByVersion("Visio", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVValidationIssue get_ItemFromID(Int32 issueID)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVValidationIssue>(this, "ItemFromID", NetOffice.VisioApi.IVValidationIssue.LateBindingApiWrapperType, issueID);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Alias for get_ItemFromID
		/// </summary>
		/// <param name="issueID">Int32 issueID</param>
		[SupportByVersion("Visio", 14,15,16), Redirect("get_ItemFromID")]
		public NetOffice.VisioApi.IVValidationIssue ItemFromID(Int32 issueID)
		{
			return get_ItemFromID(issueID);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VisioApi.IVValidationIssue>

        ICOMObject IEnumerableProvider<NetOffice.VisioApi.IVValidationIssue>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.VisioApi.IVValidationIssue>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VisioApi.IVValidationIssue>

        /// <summary>
        /// SupportByVersion Visio, 14,15,16
        /// </summary>
        [SupportByVersion("Visio", 14, 15, 16)]
        public IEnumerator<NetOffice.VisioApi.IVValidationIssue> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VisioApi.IVValidationIssue item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Visio, 14,15,16
        /// </summary>
        [SupportByVersion("Visio", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}