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
	/// DispatchInterface Trendlines 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class Trendlines : COMObject, IEnumerableProvider<NetOffice.WordApi.Trendline>
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
                    _type = typeof(Trendlines);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Trendlines(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Trendlines(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Count"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, type, order);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, type, order, period);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, type, order, period, forward);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.WordApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "Add", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Trendline this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Trendline>(this, "_Default", NetOffice.WordApi.Trendline.LateBindingApiWrapperType, index);
			}
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Trendline>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Trendline>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Trendline>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Trendline>

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Trendline> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Trendline item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}