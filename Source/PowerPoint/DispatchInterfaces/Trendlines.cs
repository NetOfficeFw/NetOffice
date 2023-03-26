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
	/// DispatchInterface Trendlines 
	/// SupportByVersion PowerPoint, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines"/> </remarks>
	[SupportByVersion("PowerPoint", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "_Default")]
	public class Trendlines : COMObject, IEnumerableProvider<NetOffice.PowerPointApi.Trendline>
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
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Count"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, type, order);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, type, order, period);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, type, order, period, forward);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Trendlines.Add"/> </remarks>
		/// <param name="type">optional NetOffice.PowerPointApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Trendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "Add", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Trendline this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Trendline>(this, "_Default", NetOffice.PowerPointApi.Trendline.LateBindingApiWrapperType, index);
			}
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PowerPointApi.Trendline>

        ICOMObject IEnumerableProvider<NetOffice.PowerPointApi.Trendline>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.PowerPointApi.Trendline>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Trendline>

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public IEnumerator<NetOffice.PowerPointApi.Trendline> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PowerPointApi.Trendline item in innerEnumerator)
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