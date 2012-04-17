using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface ContentControl 
	/// SupportByVersion Word, 12,14
	///</summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ContentControl : COMObject
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
                    _type = typeof(ContentControl);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ContentControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
				NetOffice.WordApi.Range newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool LockContentControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockContentControl", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockContentControl", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool LockContents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LockContents", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LockContents", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.XMLMapping XMLMapping
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLMapping", paramsArray);
				NetOffice.WordApi.XMLMapping newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.XMLMapping.LateBindingApiWrapperType) as NetOffice.WordApi.XMLMapping;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Enums.WdContentControlType Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdContentControlType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Type", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.ContentControlListEntries DropdownListEntries
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DropdownListEntries", paramsArray);
				NetOffice.WordApi.ContentControlListEntries newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControlListEntries.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControlListEntries;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.BuildingBlock PlaceholderText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PlaceholderText", paramsArray);
				NetOffice.WordApi.BuildingBlock newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.BuildingBlock.LateBindingApiWrapperType) as NetOffice.WordApi.BuildingBlock;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string Title
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Title", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Title", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string DateDisplayFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateDisplayFormat", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DateDisplayFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool MultiLine
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MultiLine", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MultiLine", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ParentContentControl", paramsArray);
				NetOffice.WordApi.ContentControl newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.ContentControl.LateBindingApiWrapperType) as NetOffice.WordApi.ContentControl;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool Temporary
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Temporary", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Temporary", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public bool ShowingPlaceholderText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ShowingPlaceholderText", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Enums.WdContentControlDateStorageFormat DateStorageFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateStorageFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdContentControlDateStorageFormat)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DateStorageFormat", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Enums.WdBuildingBlockTypes BuildingBlockType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildingBlockType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdBuildingBlockTypes)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BuildingBlockType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string BuildingBlockCategory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildingBlockCategory", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BuildingBlockCategory", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Enums.WdLanguageID DateDisplayLocale
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateDisplayLocale", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdLanguageID)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DateDisplayLocale", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public object DefaultTextStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTextStyle", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTextStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.Enums.WdCalendarType DateCalendarType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DateCalendarType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdCalendarType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DateCalendarType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public string Tag
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tag", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Tag", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 14)]
		public bool Checked
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Checked", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Checked", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="deleteContents">optional bool DeleteContents = false</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public void Delete(bool deleteContents)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(deleteContents);
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		/// <param name="text">optional string Text = </param>
		[SupportByVersionAttribute("Word", 12,14)]
		public void SetPlaceholderText(NetOffice.WordApi.BuildingBlock buildingBlock, NetOffice.WordApi.Range range, string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(buildingBlock, range, text);
			Invoker.Method(this, "SetPlaceholderText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public void SetPlaceholderText()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetPlaceholderText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public void SetPlaceholderText(NetOffice.WordApi.BuildingBlock buildingBlock)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(buildingBlock);
			Invoker.Method(this, "SetPlaceholderText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="buildingBlock">optional NetOffice.WordApi.BuildingBlock BuildingBlock = 0</param>
		/// <param name="range">optional NetOffice.WordApi.Range Range = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public void SetPlaceholderText(NetOffice.WordApi.BuildingBlock buildingBlock, NetOffice.WordApi.Range range)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(buildingBlock, range);
			Invoker.Method(this, "SetPlaceholderText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		public void Ungroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Ungroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersionAttribute("Word", 14)]
		public void SetCheckedSymbol(Int32 characterNumber, string font)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font);
			Invoker.Method(this, "SetCheckedSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14)]
		public void SetCheckedSymbol(Int32 characterNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber);
			Invoker.Method(this, "SetCheckedSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		/// <param name="font">optional string Font = </param>
		[SupportByVersionAttribute("Word", 14)]
		public void SetUncheckedSymbol(Int32 characterNumber, string font)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber, font);
			Invoker.Method(this, "SetUncheckedSymbol", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 14
		/// </summary>
		/// <param name="characterNumber">Int32 CharacterNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 14)]
		public void SetUncheckedSymbol(Int32 characterNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(characterNumber);
			Invoker.Method(this, "SetUncheckedSymbol", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}