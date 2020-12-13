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
	/// DispatchInterface ProtectedViewWindows 
	/// SupportByVersion Word, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows"/> </remarks>
	[SupportByVersion("Word", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class ProtectedViewWindows : COMObject, IEnumerableProvider<NetOffice.WordApi.ProtectedViewWindow>
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
                    _type = typeof(ProtectedViewWindows);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ProtectedViewWindows(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ProtectedViewWindows(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProtectedViewWindows(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Application"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Creator"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Count"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
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
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.ProtectedViewWindow this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Item", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible, object openAndRepair)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Open", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, new object[]{ fileName, addToRecentFiles, passwordDocument, visible, openAndRepair });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Open", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Open", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, fileName, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Open", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, fileName, addToRecentFiles, passwordDocument);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.ProtectedViewWindows.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ProtectedViewWindow Open(object fileName, object addToRecentFiles, object passwordDocument, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ProtectedViewWindow>(this, "Open", NetOffice.WordApi.ProtectedViewWindow.LateBindingApiWrapperType, fileName, addToRecentFiles, passwordDocument, visible);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.ProtectedViewWindow>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.ProtectedViewWindow>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.ProtectedViewWindow>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.ProtectedViewWindow>

        /// <summary>
        /// SupportByVersion Word, 14,15,16
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.ProtectedViewWindow> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.ProtectedViewWindow item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}