using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface SeriesCollection 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821223.aspx </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class SeriesCollection : COMObject, IEnumerableProvider<NetOffice.WordApi.Series>
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
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839574.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834899.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836348.aspx </remarks>
		[SupportByVersion("Word", 14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845293.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192549.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.WordApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		/// <param name="replace">optional object replace</param>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "Add", NetOffice.WordApi.Series.LateBindingApiWrapperType, new object[]{ source, rowcol, seriesLabels, categoryLabels, replace });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192549.aspx </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series Add(object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "Add", NetOffice.WordApi.Series.LateBindingApiWrapperType, source);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192549.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.WordApi.Enums.XlRowCol Rowcol = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series Add(object source, object rowcol)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "Add", NetOffice.WordApi.Series.LateBindingApiWrapperType, source, rowcol);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192549.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.WordApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series Add(object source, object rowcol, object seriesLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "Add", NetOffice.WordApi.Series.LateBindingApiWrapperType, source, rowcol, seriesLabels);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192549.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional NetOffice.WordApi.Enums.XlRowCol Rowcol = 2</param>
		/// <param name="seriesLabels">optional object seriesLabels</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series Add(object source, object rowcol, object seriesLabels, object categoryLabels)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "Add", NetOffice.WordApi.Series.LateBindingApiWrapperType, source, rowcol, seriesLabels, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197201.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		/// <param name="categoryLabels">optional object categoryLabels</param>
		[SupportByVersion("Word", 14,15,16)]
		public object Extend(object source, object rowcol, object categoryLabels)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol, categoryLabels);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197201.aspx </remarks>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Extend(object source)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197201.aspx </remarks>
		/// <param name="source">object source</param>
		/// <param name="rowcol">optional object rowcol</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public object Extend(object source, object rowcol)
		{
			return Factory.ExecuteVariantMethodGet(this, "Extend", source, rowcol);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840883.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Series NewSeries()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "NewSeries", NetOffice.WordApi.Series.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Series this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Series>(this, "_Default", NetOffice.WordApi.Series.LateBindingApiWrapperType, index);
			}
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Series>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Series>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Series>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Series>

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Series> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Series item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, false);
		}

		#endregion

		#pragma warning restore
	}
}