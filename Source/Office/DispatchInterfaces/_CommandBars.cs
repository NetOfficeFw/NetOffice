using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface _CommandBars 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _CommandBars : _IMsoDispObj ,IEnumerable<NetOffice.OfficeApi.CommandBar>
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_CommandBars);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CommandBars(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862425.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl ActionControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActionControl", paramsArray);
				NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863075.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar ActiveMenuBar
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveMenuBar", paramsArray);
				NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860520.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863160.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public bool DisplayTooltips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayTooltips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayTooltips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864956.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public bool DisplayKeysInTooltips
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayKeysInTooltips", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayKeysInTooltips", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.OfficeApi.CommandBar this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864068.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public bool LargeButtons
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LargeButtons", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LargeButtons", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864076.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoMenuAnimation MenuAnimationStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MenuAnimationStyle", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OfficeApi.Enums.MsoMenuAnimation)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MenuAnimationStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862543.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="ids">Int32 ids</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_IdsString(Int32 ids, out string pbstrName)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(ids, pbstrName);
			object returnItem = Invoker.PropertyGet(this, "IdsString", paramsArray);
			pbstrName = (string)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IdsString
		/// </summary>
		/// <param name="ids">Int32 ids</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public Int32 IdsString(Int32 ids, out string pbstrName)
		{
			return get_IdsString(ids, out pbstrName);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="tmc">Int32 tmc</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_TmcGetName(Int32 tmc, out string pbstrName)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(tmc, pbstrName);
			object returnItem = Invoker.PropertyGet(this, "TmcGetName", paramsArray);
			pbstrName = (string)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_TmcGetName
		/// </summary>
		/// <param name="tmc">Int32 tmc</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public Int32 TmcGetName(Int32 tmc, out string pbstrName)
		{
			return get_TmcGetName(tmc, out pbstrName);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860590.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public bool AdaptiveMenus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AdaptiveMenus", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AdaptiveMenus", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860823.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public bool DisplayFonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayFonts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayFonts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864631.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool DisableCustomize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisableCustomize", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisableCustomize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863405.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 10,11,12,14,15,16)]
		public bool DisableAskAQuestionDropdown
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisableAskAQuestionDropdown", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisableAskAQuestionDropdown", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="position">optional object Position</param>
		/// <param name="menuBar">optional object MenuBar</param>
		/// <param name="temporary">optional object Temporary</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar, object temporary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, position, menuBar, temporary);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="position">optional object Position</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, position);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861773.aspx
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="position">optional object Position</param>
		/// <param name="menuBar">optional object MenuBar</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, position, menuBar);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		/// <param name="tag">optional object Tag</param>
		/// <param name="visible">optional object Visible</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id, tag, visible);
			object returnItem = Invoker.MethodReturn(this, "FindControl", paramsArray);
			NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindControl", paramsArray);
			NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "FindControl", paramsArray);
			NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id);
			object returnItem = Invoker.MethodReturn(this, "FindControl", paramsArray);
			NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860267.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		/// <param name="tag">optional object Tag</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id, tag);
			object returnItem = Invoker.MethodReturn(this, "FindControl", paramsArray);
			NetOffice.OfficeApi.CommandBarControl newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OfficeApi.CommandBarControl;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861062.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public void ReleaseFocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReleaseFocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		/// <param name="tag">optional object Tag</param>
		/// <param name="visible">optional object Visible</param>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id, tag, visible);
			object returnItem = Invoker.MethodReturn(this, "FindControls", paramsArray);
			NetOffice.OfficeApi.CommandBarControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBarControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "FindControls", paramsArray);
			NetOffice.OfficeApi.CommandBarControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBarControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type);
			object returnItem = Invoker.MethodReturn(this, "FindControls", paramsArray);
			NetOffice.OfficeApi.CommandBarControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBarControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id);
			object returnItem = Invoker.MethodReturn(this, "FindControls", paramsArray);
			NetOffice.OfficeApi.CommandBarControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBarControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862556.aspx
		/// </summary>
		/// <param name="type">optional object Type</param>
		/// <param name="id">optional object Id</param>
		/// <param name="tag">optional object Tag</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, id, tag);
			object returnItem = Invoker.MethodReturn(this, "FindControls", paramsArray);
			NetOffice.OfficeApi.CommandBarControls newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBarControls;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tbidOrName">optional object TbidOrName</param>
		/// <param name="position">optional object Position</param>
		/// <param name="menuBar">optional object MenuBar</param>
		/// <param name="temporary">optional object Temporary</param>
		/// <param name="tbtrProtection">optional object TbtrProtection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary, object tbtrProtection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tbidOrName, position, menuBar, temporary, tbtrProtection);
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tbidOrName">optional object TbidOrName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tbidOrName);
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tbidOrName">optional object TbidOrName</param>
		/// <param name="position">optional object Position</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tbidOrName, position);
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tbidOrName">optional object TbidOrName</param>
		/// <param name="position">optional object Position</param>
		/// <param name="menuBar">optional object MenuBar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tbidOrName, position, menuBar);
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tbidOrName">optional object TbidOrName</param>
		/// <param name="position">optional object Position</param>
		/// <param name="menuBar">optional object MenuBar</param>
		/// <param name="temporary">optional object Temporary</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tbidOrName, position, menuBar, temporary);
			object returnItem = Invoker.MethodReturn(this, "AddEx", paramsArray);
			NetOffice.OfficeApi.CommandBar newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType) as NetOffice.OfficeApi.CommandBar;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862419.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public void ExecuteMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			Invoker.Method(this, "ExecuteMso", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862202.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool GetEnabledMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetEnabledMso", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863712.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool GetVisibleMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetVisibleMso", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863149.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public bool GetPressedMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetPressedMso", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860585.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public string GetLabelMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetLabelMso", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860790.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public string GetScreentipMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetScreentipMso", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864975.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public string GetSupertipMso(string idMso)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso);
			object returnItem = Invoker.MethodReturn(this, "GetSupertipMso", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861156.aspx
		/// </summary>
		/// <param name="idMso">string idMso</param>
		/// <param name="width">Int32 Width</param>
		/// <param name="height">Int32 Height</param>
		[SupportByVersionAttribute("Office", 12,14,15,16)]
		public stdole.Picture GetImageMso(string idMso, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso, width, height);
			object returnItem = Invoker.MethodReturn(this, "GetImageMso", paramsArray);
			stdole.Picture newObject = Factory.CreateObjectFromComProxy(this, returnItem) as stdole.Picture;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863478.aspx
		/// </summary>
		/// <param name="hwnd">Int32 hwnd</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public void CommitRenderingTransaction(Int32 hwnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hwnd);
			Invoker.Method(this, "CommitRenderingTransaction", paramsArray);
		}

		#endregion

       #region IEnumerable<NetOffice.OfficeApi.CommandBar> Member
        
        /// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.OfficeApi.CommandBar> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.OfficeApi.CommandBar item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Office, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}