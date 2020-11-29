﻿using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface Conflicts 
	/// SupportByVersion Outlook, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts"/> </remarks>
	[SupportByVersion("Outlook", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Custom), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Conflicts : COMObject, IEnumerableProvider<NetOffice.OutlookApi.Conflict>
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
                    _type = typeof(Conflicts);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Conflicts(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Conflicts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Conflicts(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.Application"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.Class"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.Session"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.Parent"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.Count"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
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
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OutlookApi.Conflict this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Conflict>(this, "Item", NetOffice.OutlookApi.Conflict.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.GetFirst"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflict GetFirst()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Conflict>(this, "GetFirst", NetOffice.OutlookApi.Conflict.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.GetLast"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflict GetLast()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Conflict>(this, "GetLast", NetOffice.OutlookApi.Conflict.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.GetNext"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflict GetNext()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Conflict>(this, "GetNext", NetOffice.OutlookApi.Conflict.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.Conflicts.GetPrevious"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflict GetPrevious()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.Conflict>(this, "GetPrevious", NetOffice.OutlookApi.Conflict.LateBindingApiWrapperType);
		}

        #endregion
        
        #region IEnumerableProvider<NetOffice.OutlookApi.Conflict>

        ICOMObject IEnumerableProvider<NetOffice.OutlookApi.Conflict>.GetComObjectEnumerator(ICOMObject parent)
        {
            return this;
        }

        IEnumerable IEnumerableProvider<NetOffice.OutlookApi.Conflict>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OutlookApi.Conflict item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable<NetOffice.OutlookApi.Conflict>

        /// <summary>
        /// SupportByVersion Outlook, 11,12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 11, 12, 14, 15, 16)]
        [CustomEnumerator]
        public IEnumerator<NetOffice.OutlookApi.Conflict> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OutlookApi.Conflict item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Outlook, 11,12,14,15,16
        /// This is a custom enumerator from NetOffice
        /// </summary>
        [SupportByVersion("Outlook", 11, 12, 14, 15, 16)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i + 1];

            foreach (object item in enumeratorObjects)
                yield return item;
        }

        #endregion

        #pragma warning restore
    }
}