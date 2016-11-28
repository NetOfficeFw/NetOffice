using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// Interface IMenuItems 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IMenuItems : COMObject ,IEnumerable<object>
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
                    _type = typeof(IMenuItems);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMenuItems(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMenuItems(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="before">optional object Before</param>
		/// <param name="restore">optional object Restore</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpFile">optional object HelpFile</param>
		/// <param name="helpContextID">optional object HelpContextID</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar, object helpFile, object helpContextID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey, before, restore, statusBar, helpFile, helpContextID);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey, before);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="before">optional object Before</param>
		/// <param name="restore">optional object Restore</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey, before, restore);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="before">optional object Before</param>
		/// <param name="restore">optional object Restore</param>
		/// <param name="statusBar">optional object StatusBar</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey, before, restore, statusBar);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="onAction">optional object OnAction</param>
		/// <param name="shortcutKey">optional object ShortcutKey</param>
		/// <param name="before">optional object Before</param>
		/// <param name="restore">optional object Restore</param>
		/// <param name="statusBar">optional object StatusBar</param>
		/// <param name="helpFile">optional object HelpFile</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.MenuItem Add(string caption, object onAction, object shortcutKey, object before, object restore, object statusBar, object helpFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, onAction, shortcutKey, before, restore, statusBar, helpFile);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.MenuItem newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.MenuItem.LateBindingApiWrapperType) as NetOffice.ExcelApi.MenuItem;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="before">optional object Before</param>
		/// <param name="restore">optional object Restore</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption, object before, object restore)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, before, restore);
			object returnItem = Invoker.MethodReturn(this, "AddMenu", paramsArray);
			NetOffice.ExcelApi.Menu newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Menu.LateBindingApiWrapperType) as NetOffice.ExcelApi.Menu;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption);
			object returnItem = Invoker.MethodReturn(this, "AddMenu", paramsArray);
			NetOffice.ExcelApi.Menu newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Menu.LateBindingApiWrapperType) as NetOffice.ExcelApi.Menu;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="caption">string Caption</param>
		/// <param name="before">optional object Before</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Menu AddMenu(string caption, object before)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(caption, before);
			object returnItem = Invoker.MethodReturn(this, "AddMenu", paramsArray);
			NetOffice.ExcelApi.Menu newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Menu.LateBindingApiWrapperType) as NetOffice.ExcelApi.Menu;
			return newObject;
		}

		#endregion

       #region IEnumerable<object> Member
        
        /// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
       public IEnumerator<object> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (object item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}