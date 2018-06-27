using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface SeriesCollection 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    public class SeriesCollection : COMObject, NetOffice.OfficeApi.SeriesCollection
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.SeriesCollection);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SeriesCollection() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.IMsoSeries this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoSeries>(this, "_Default", typeof(NetOffice.OfficeApi.IMsoSeries), index);
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels, object replace)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", typeof(NetOffice.OfficeApi.IMsoSeries), new object[] { source, rowcol, seriesLabels, categoryLabels, replace });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries Add(object source)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", typeof(NetOffice.OfficeApi.IMsoSeries), source);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", typeof(NetOffice.OfficeApi.IMsoSeries), source, rowcol);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", typeof(NetOffice.OfficeApi.IMsoSeries), source, rowcol, seriesLabels);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries Add(object source, object rowcol, object seriesLabels, object categoryLabels)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "Add", typeof(NetOffice.OfficeApi.IMsoSeries), source, rowcol, seriesLabels, categoryLabels);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional object rowcol</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Extend(object source, object rowcol, object categoryLabels)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Extend", source, rowcol, categoryLabels);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Extend(object source)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Extend", source);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">object source</param>
        /// <param name="rowcol">optional object rowcol</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Extend(object source, object rowcol)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Extend", source, rowcol);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="replace">optional object replace</param>
        /// <param name="newSeries">optional object newSeries</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace, object newSeries)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste", new object[] { rowcol, seriesLabels, categoryLabels, replace, newSeries });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste(object rowcol)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste", rowcol);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste(object rowcol, object seriesLabels)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste(object rowcol, object seriesLabels, object categoryLabels)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels, categoryLabels);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rowcol">optional NetOffice.OfficeApi.Enums.XlRowCol Rowcol = 2</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="replace">optional object replace</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste(object rowcol, object seriesLabels, object categoryLabels, object replace)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste", rowcol, seriesLabels, categoryLabels, replace);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeries NewSeries()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoSeries>(this, "NewSeries", typeof(NetOffice.OfficeApi.IMsoSeries));
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.IMsoSeries>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
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
        public virtual IEnumerator<NetOffice.OfficeApi.IMsoSeries> GetEnumerator()
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
