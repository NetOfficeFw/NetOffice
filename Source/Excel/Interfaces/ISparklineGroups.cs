using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISparklineGroups 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class ISparklineGroups : COMObject, IEnumerableProvider<NetOffice.ExcelApi.SparklineGroup>
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
                    _type = typeof(ISparklineGroups);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISparklineGroups(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISparklineGroups(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroups(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.SparklineGroup this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SparklineGroup>(this, "_Default", NetOffice.ExcelApi.SparklineGroup.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlSparkType type</param>
		/// <param name="sourceData">string sourceData</param>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SparklineGroup Add(NetOffice.ExcelApi.Enums.XlSparkType type, string sourceData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SparklineGroup>(this, "Add", NetOffice.ExcelApi.SparklineGroup.LateBindingApiWrapperType, type, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Clear()
		{
			return Factory.ExecuteInt32MethodGet(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ClearGroups()
		{
			return Factory.ExecuteInt32MethodGet(this, "ClearGroups");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="location">NetOffice.ExcelApi.Range location</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Group(NetOffice.ExcelApi.Range location)
		{
			return Factory.ExecuteInt32MethodGet(this, "Group", location);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Ungroup()
		{
			return Factory.ExecuteInt32MethodGet(this, "Ungroup");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.SparklineGroup>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.SparklineGroup>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.SparklineGroup>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.SparklineGroup>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.SparklineGroup> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.SparklineGroup item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}