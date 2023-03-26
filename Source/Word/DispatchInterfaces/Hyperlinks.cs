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
	/// DispatchInterface Hyperlinks 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.hyperlinks"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Hyperlinks : COMObject, IEnumerableProvider<NetOffice.WordApi.Hyperlink>
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
                    _type = typeof(Hyperlinks);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Hyperlinks(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Hyperlinks(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Hyperlink this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Item", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor, object address, object subAddress)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "_Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor, address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">object anchor</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "_Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink _Add(object anchor, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "_Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="screenTip">optional object screenTip</param>
		/// <param name="textToDisplay">optional object textToDisplay</param>
		/// <param name="target">optional object target</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip, object textToDisplay, object target)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, new object[]{ anchor, address, subAddress, screenTip, textToDisplay, target });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor, address);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor, address, subAddress);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="screenTip">optional object screenTip</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, anchor, address, subAddress, screenTip);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Hyperlinks.Add"/> </remarks>
		/// <param name="anchor">object anchor</param>
		/// <param name="address">optional object address</param>
		/// <param name="subAddress">optional object subAddress</param>
		/// <param name="screenTip">optional object screenTip</param>
		/// <param name="textToDisplay">optional object textToDisplay</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlink Add(object anchor, object address, object subAddress, object screenTip, object textToDisplay)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Hyperlink>(this, "Add", NetOffice.WordApi.Hyperlink.LateBindingApiWrapperType, new object[]{ anchor, address, subAddress, screenTip, textToDisplay });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Hyperlink>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Hyperlink>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Hyperlink>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Hyperlink>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Hyperlink> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Hyperlink item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}