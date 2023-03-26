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
	/// DispatchInterface MenuItems 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class MenuItems : COMObject, IEnumerableProvider<object>
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
                    _type = typeof(MenuItems);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public MenuItems(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MenuItems(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MenuItems(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public object this[object index]
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "_Default", index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="before">optional object before</param>
		/// <param name="restore">optional object restore</param>
		/// <param name="statusBar">optional object statusBar</param>
		/// <param name="helpFile">optional object helpFile</param>
		/// <param name="helpContextID">optional object helpContextID</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar, object helpFile, object helpContextID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, new object[]{ caption, onAction, shortcutKey, before, restore, statusBar, helpFile, helpContextID });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, caption);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, caption, onAction);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, caption, onAction, shortcutKey);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, caption, onAction, shortcutKey, before);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="before">optional object before</param>
		/// <param name="restore">optional object restore</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, new object[]{ caption, onAction, shortcutKey, before, restore });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="before">optional object before</param>
		/// <param name="restore">optional object restore</param>
		/// <param name="statusBar">optional object statusBar</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, new object[]{ caption, onAction, shortcutKey, before, restore, statusBar });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="onAction">optional object onAction</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="before">optional object before</param>
		/// <param name="restore">optional object restore</param>
		/// <param name="statusBar">optional object statusBar</param>
		/// <param name="helpFile">optional object helpFile</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar, object helpFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.MenuItem>(this, "Add", NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType, new object[]{ caption, onAction, shortcutKey, before, restore, statusBar, helpFile });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="before">optional object before</param>
		/// <param name="restore">optional object restore</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption, object before, object restore)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Menu>(this, "AddMenu", NetOffice.ExcelApi.Menu.LateBindingApiWrapperType, caption, before, restore);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Menu>(this, "AddMenu", NetOffice.ExcelApi.Menu.LateBindingApiWrapperType, caption);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="caption">string caption</param>
		/// <param name="before">optional object before</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption, object before)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Menu>(this, "AddMenu", NetOffice.ExcelApi.Menu.LateBindingApiWrapperType, caption, before);
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}