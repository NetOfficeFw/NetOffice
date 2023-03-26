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
	/// Trendlines
	/// </summary>
	[SyntaxBypass]
 	public class Trendlines_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Trendlines_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Trendlines_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Trendlines_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// Interface Trendlines 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Trendlines : COMObject, IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>
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
		/// <param name="index">optional object index</param>
		[SupportByVersion("Office", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.IMsoTrendline this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoTrendline>(this, "_Default", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared, name });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, type, order);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, type, order, period);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, type, order, period, forward);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlTrendlineType Type = -4132</param>
		/// <param name="order">optional object order</param>
		/// <param name="period">optional object period</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="backward">optional object backward</param>
		/// <param name="intercept">optional object intercept</param>
		/// <param name="displayEquation">optional object displayEquation</param>
		/// <param name="displayRSquared">optional object displayRSquared</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.IMsoTrendline Add(object type, object order, object period, object forward, object backward, object intercept, object displayEquation, object displayRSquared)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.IMsoTrendline>(this, "Add", NetOffice.OfficeApi.IMsoTrendline.LateBindingApiWrapperType, new object[]{ type, order, period, forward, backward, intercept, displayEquation, displayRSquared });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.IMsoTrendline>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.IMsoTrendline>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.IMsoTrendline> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.IMsoTrendline item in innerEnumerator)
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