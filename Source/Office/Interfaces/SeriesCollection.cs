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
	/// Interface SeriesCollection 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class SeriesCollection : COMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
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

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.IMsoSeries this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoSeries>(this, "_Default", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, new object[]{ source, rowcol, seriesLabels, categoryLabels, replace });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries Add(object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, source);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, source, rowcol);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, source, rowcol, seriesLabels);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType, source, rowcol, seriesLabels, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Extend(object source, object rowcol, object categoryLabels)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Extend(object source)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Extend(object source, object rowcol)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="newSeries">optional object newSeries</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace, object newSeries)
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste", new object[]{ rowcol, seriesLabels, categoryLabels, replace, newSeries });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste()
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste(object rowcol)
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste", rowcol);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste(object rowcol, object seriesLabels)
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste(object rowcol, object seriesLabels, object categoryLabels)
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace)
		{
			return Factory.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels, categoryLabels, replace);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoSeries NewSeries()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "NewSeries", NetOffice.OfficeApi.IMsoSeries.LateBindingApiWrapperType);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoSeries>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.IMsoSeries> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.IMsoSeries item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}