using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface _CommandBars 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class _CommandBars : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.CommandBar>
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
                    _type = typeof(_CommandBars);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CommandBars(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CommandBars(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.ActionControl"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OfficeApi.CommandBarControl ActionControl
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OfficeApi.CommandBarControl>(this, "ActionControl");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.ActiveMenuBar"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar ActiveMenuBar
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "ActiveMenuBar", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Count"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.DisplayTooltips"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool DisplayTooltips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayTooltips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayTooltips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.DisplayKeysInTooltips"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool DisplayKeysInTooltips
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayKeysInTooltips");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayKeysInTooltips", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.CommandBar this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBar>(this, "Item", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.LargeButtons"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool LargeButtons
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LargeButtons");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LargeButtons", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.MenuAnimationStyle"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoMenuAnimation MenuAnimationStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoMenuAnimation>(this, "MenuAnimationStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MenuAnimationStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Parent"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="ids">Int32 ids</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_IdsString(Int32 ids, out string pbstrName)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(ids, pbstrName);
			object returnItem = Invoker.PropertyGet(this, "IdsString", paramsArray, modifiers);
			pbstrName = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_IdsString
		/// </summary>
		/// <param name="ids">Int32 ids</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_IdsString")]
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_TmcGetName(Int32 tmc, out string pbstrName)
		{		
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrName = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(tmc, pbstrName);
			object returnItem = Invoker.PropertyGet(this, "TmcGetName", paramsArray, modifiers);
			pbstrName = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_TmcGetName
		/// </summary>
		/// <param name="tmc">Int32 tmc</param>
		/// <param name="pbstrName">string pbstrName</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), Redirect("get_TmcGetName")]
		public Int32 TmcGetName(Int32 tmc, out string pbstrName)
		{
			return get_TmcGetName(tmc, out pbstrName);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.AdaptiveMenus"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool AdaptiveMenus
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AdaptiveMenus");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AdaptiveMenus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.DisplayFonts"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public bool DisplayFonts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayFonts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayFonts", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.DisableCustomize"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool DisableCustomize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableCustomize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableCustomize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.DisableAskAQuestionDropdown"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool DisableAskAQuestionDropdown
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableAskAQuestionDropdown");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableAskAQuestionDropdown", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="position">optional object position</param>
		/// <param name="menuBar">optional object menuBar</param>
		/// <param name="temporary">optional object temporary</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar, object temporary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, name, position, menuBar, temporary);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="position">optional object position</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, name, position);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.Add"/> </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="position">optional object position</param>
		/// <param name="menuBar">optional object menuBar</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar Add(object name, object position, object menuBar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "Add", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, name, position, menuBar);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControl"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag, object visible)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag, visible);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControl"/> </remarks>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControl"/> </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControl"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControl"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControl FindControl(object type, object id, object tag)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi.CommandBarControl>(this, "FindControl", type, id, tag);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.ReleaseFocus"/> </remarks>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void ReleaseFocus()
		{
			 Factory.ExecuteMethod(this, "ReleaseFocus");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControls"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType, type, id, tag, visible);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControls"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControls"/> </remarks>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControls"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType, type, id);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.FindControls"/> </remarks>
		/// <param name="type">optional object type</param>
		/// <param name="id">optional object id</param>
		/// <param name="tag">optional object tag</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBarControls FindControls(object type, object id, object tag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBarControls>(this, "FindControls", NetOffice.OfficeApi.CommandBarControls.LateBindingApiWrapperType, type, id, tag);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tbidOrName">optional object tbidOrName</param>
		/// <param name="position">optional object position</param>
		/// <param name="menuBar">optional object menuBar</param>
		/// <param name="temporary">optional object temporary</param>
		/// <param name="tbtrProtection">optional object tbtrProtection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary, object tbtrProtection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, new object[]{ tbidOrName, position, menuBar, temporary, tbtrProtection });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tbidOrName">optional object tbidOrName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, tbidOrName);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tbidOrName">optional object tbidOrName</param>
		/// <param name="position">optional object position</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, tbidOrName, position);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tbidOrName">optional object tbidOrName</param>
		/// <param name="position">optional object position</param>
		/// <param name="menuBar">optional object menuBar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, tbidOrName, position, menuBar);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tbidOrName">optional object tbidOrName</param>
		/// <param name="position">optional object position</param>
		/// <param name="menuBar">optional object menuBar</param>
		/// <param name="temporary">optional object temporary</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.CommandBar AddEx(object tbidOrName, object position, object menuBar, object temporary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.CommandBar>(this, "AddEx", NetOffice.OfficeApi.CommandBar.LateBindingApiWrapperType, tbidOrName, position, menuBar, temporary);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.ExecuteMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void ExecuteMso(string idMso)
		{
			 Factory.ExecuteMethod(this, "ExecuteMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetEnabledMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool GetEnabledMso(string idMso)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetEnabledMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetVisibleMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool GetVisibleMso(string idMso)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetVisibleMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetPressedMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public bool GetPressedMso(string idMso)
		{
			return Factory.ExecuteBoolMethodGet(this, "GetPressedMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetLabelMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public string GetLabelMso(string idMso)
		{
			return Factory.ExecuteStringMethodGet(this, "GetLabelMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetScreentipMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public string GetScreentipMso(string idMso)
		{
			return Factory.ExecuteStringMethodGet(this, "GetScreentipMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetSupertipMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public string GetSupertipMso(string idMso)
		{
			return Factory.ExecuteStringMethodGet(this, "GetSupertipMso", idMso);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.GetImageMso"/> </remarks>
		/// <param name="idMso">string idMso</param>
		/// <param name="width">Int32 width</param>
		/// <param name="height">Int32 height</param>
		[SupportByVersion("Office", 12,14,15,16), NativeResult]
		public stdole.Picture GetImageMso(string idMso, Int32 width, Int32 height)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(idMso, width, height);
			object returnItem = Invoker.MethodReturn(this, "GetImageMso", paramsArray);
            return returnItem as stdole.Picture;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.CommandBars.CommitRenderingTransaction"/> </remarks>
		/// <param name="hwnd">Int32 hwnd</param>
		[SupportByVersion("Office", 14,15,16)]
		public void CommitRenderingTransaction(Int32 hwnd)
		{
			 Factory.ExecuteMethod(this, "CommitRenderingTransaction", hwnd);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.CommandBar>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.CommandBar>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.CommandBar>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.CommandBar>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.CommandBar> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.CommandBar item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}