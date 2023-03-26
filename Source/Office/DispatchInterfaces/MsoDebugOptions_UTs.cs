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
	/// DispatchInterface MsoDebugOptions_UTs 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class MsoDebugOptions_UTs : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.MsoDebugOptions_UT>
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
                    _type = typeof(MsoDebugOptions_UTs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MsoDebugOptions_UTs(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MsoDebugOptions_UTs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoDebugOptions_UTs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.MsoDebugOptions_UT this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.MsoDebugOptions_UT>(this, "Item", NetOffice.OfficeApi.MsoDebugOptions_UT.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCollectionName">string bstrCollectionName</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.MsoDebugOptions_UTs GetUnitTestsInCollection(string bstrCollectionName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.MsoDebugOptions_UTs>(this, "GetUnitTestsInCollection", NetOffice.OfficeApi.MsoDebugOptions_UTs.LateBindingApiWrapperType, bstrCollectionName);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCollectionName">string bstrCollectionName</param>
		/// <param name="bstrUnitTestName">string bstrUnitTestName</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.MsoDebugOptions_UT GetUnitTest(string bstrCollectionName, string bstrUnitTestName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.MsoDebugOptions_UT>(this, "GetUnitTest", NetOffice.OfficeApi.MsoDebugOptions_UT.LateBindingApiWrapperType, bstrCollectionName, bstrUnitTestName);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrCollectionName">string bstrCollectionName</param>
		/// <param name="bstrUnitTestNameFilter">string bstrUnitTestNameFilter</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.MsoDebugOptions_UTs GetMatchingUnitTestsInCollection(string bstrCollectionName, string bstrUnitTestNameFilter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.MsoDebugOptions_UTs>(this, "GetMatchingUnitTestsInCollection", NetOffice.OfficeApi.MsoDebugOptions_UTs.LateBindingApiWrapperType, bstrCollectionName, bstrUnitTestNameFilter);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.MsoDebugOptions_UT>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.MsoDebugOptions_UT>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.MsoDebugOptions_UT>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.MsoDebugOptions_UT>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.MsoDebugOptions_UT> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.MsoDebugOptions_UT item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}