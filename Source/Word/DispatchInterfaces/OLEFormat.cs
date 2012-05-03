using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface OLEFormat 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OLEFormat : COMObject
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
                    _type = typeof(OLEFormat);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OLEFormat(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ClassType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClassType", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ClassType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public bool DisplayAsIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayAsIcon", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayAsIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string IconName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IconName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IconName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string IconPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IconPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 IconIndex
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IconIndex", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IconIndex", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string IconLabel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IconLabel", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IconLabel", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string Label
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Label", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public object Object
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Object", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public string ProgID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProgID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		public bool PreserveFormattingOnUpdate
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreserveFormattingOnUpdate", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PreserveFormattingOnUpdate", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Activate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Activate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Edit()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Edit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Open()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="verbIndex">optional object VerbIndex</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void DoVerb(object verbIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(verbIndex);
			Invoker.Method(this, "DoVerb", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void DoVerb()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "DoVerb", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconIndex">optional object IconIndex</param>
		/// <param name="iconLabel">optional object IconLabel</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, displayAsIcon, iconFileName, iconIndex, iconLabel);
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo(object classType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType);
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo(object classType, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, displayAsIcon);
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, displayAsIcon, iconFileName);
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">optional object ClassType</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		/// <param name="iconFileName">optional object IconFileName</param>
		/// <param name="iconIndex">optional object IconIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ConvertTo(object classType, object displayAsIcon, object iconFileName, object iconIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType, displayAsIcon, iconFileName, iconIndex);
			Invoker.Method(this, "ConvertTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="classType">string ClassType</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void ActivateAs(string classType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classType);
			Invoker.Method(this, "ActivateAs", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}