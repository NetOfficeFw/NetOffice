using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface SeriesCollection 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class SeriesCollection : COMObject, IEnumerableProvider<NetOffice.PowerPointApi.Series>
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
                    _type = typeof(SeriesCollection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SeriesCollection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SeriesCollection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SeriesCollection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Parent"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Count"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Creator"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Application"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Extend"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Extend(object source, object rowcol, object categoryLabels)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Extend"/> </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Extend(object source)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Extend"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public object Extend(object source, object rowcol)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.NewSeries"/> </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series NewSeries()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "NewSeries", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Add"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "Add", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, new object[]{ source, rowcol, seriesLabels, categoryLabels, replace });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Add"/> </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series Add(object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "Add", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, source);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Add"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "Add", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, source, rowcol);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Add"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "Add", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, source, rowcol, seriesLabels);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.SeriesCollection.Add"/> </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.PowerPointApi.Enums.XlRowCol Rowcol = -4105</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "Add", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, source, rowcol, seriesLabels, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Series this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Series>(this, "_Default", NetOffice.PowerPointApi.Series.LateBindingApiWrapperType, index);
			}
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PowerPointApi.Series>

        ICOMObject IEnumerableProvider<NetOffice.PowerPointApi.Series>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.PowerPointApi.Series>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Series>

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public IEnumerator<NetOffice.PowerPointApi.Series> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PowerPointApi.Series item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}