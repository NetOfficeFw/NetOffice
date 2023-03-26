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
	/// Interface ISparklineGroup 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class ISparklineGroup : COMObject, IEnumerableProvider<NetOffice.ExcelApi.Sparkline>
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
                    _type = typeof(ISparklineGroup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISparklineGroup(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISparklineGroup(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISparklineGroup(string progId) : base(progId)
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
		public NetOffice.ExcelApi.Sparkline this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Sparkline>(this, "Item", NetOffice.ExcelApi.Sparkline.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Range Location
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Location", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Location", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string SourceData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string DateRange
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DateRange");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DateRange", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSparkType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSparkType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.FormatColor SeriesColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.FormatColor>(this, "SeriesColor", NetOffice.ExcelApi.FormatColor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SparkPoints Points
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SparkPoints>(this, "Points", NetOffice.ExcelApi.SparkPoints.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.SparkAxes Axes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SparkAxes>(this, "Axes", NetOffice.ExcelApi.SparkAxes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool DisplayHidden
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayHidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public object LineWeight
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LineWeight");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "LineWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSparklineRowCol PlotBy
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSparklineRowCol>(this, "PlotBy");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PlotBy", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="location">NetOffice.ExcelApi.Range location</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ModifyLocation(NetOffice.ExcelApi.Range location)
		{
			return Factory.ExecuteInt32MethodGet(this, "ModifyLocation", location);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="sourceData">string sourceData</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ModifySourceData(string sourceData)
		{
			return Factory.ExecuteInt32MethodGet(this, "ModifySourceData", sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="location">NetOffice.ExcelApi.Range location</param>
		/// <param name="sourceData">string sourceData</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Modify(NetOffice.ExcelApi.Range location, string sourceData)
		{
			return Factory.ExecuteInt32MethodGet(this, "Modify", location, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="dateRange">string dateRange</param>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 ModifyDateRange(string dateRange)
		{
			return Factory.ExecuteInt32MethodGet(this, "ModifyDateRange", dateRange);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Sparkline>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Sparkline>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Sparkline>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Sparkline>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.Sparkline> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Sparkline item in innerEnumerator)
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