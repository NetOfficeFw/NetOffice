using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// LPVISIODOCUMENT
	///</summary>
	public class LPVISIODOCUMENT_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIODOCUMENT_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_LeftMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "LeftMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_LeftMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "LeftMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_LeftMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double LeftMargin(object unitsNameOrCode)
		{
			return get_LeftMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_RightMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "RightMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_RightMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "RightMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RightMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double RightMargin(object unitsNameOrCode)
		{
			return get_RightMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_TopMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "TopMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_TopMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "TopMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_TopMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double TopMargin(object unitsNameOrCode)
		{
			return get_TopMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_BottomMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "BottomMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_BottomMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "BottomMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_BottomMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double BottomMargin(object unitsNameOrCode)
		{
			return get_BottomMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="bstrPassword">optional object bstrPassword</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.Enums.VisProtection get_Protection(object bstrPassword)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(bstrPassword);
			object returnItem = Invoker.PropertyGet(this, "Protection", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.VisioApi.Enums.VisProtection)intReturnItem;
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrPassword">optional object bstrPassword</param>
        /// <param name="value">optional VisProtection value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Protection(object bstrPassword, NetOffice.VisioApi.Enums.VisProtection value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrPassword);
			Invoker.PropertySet(this, "Protection", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Protection
		/// </summary>
		/// <param name="bstrPassword">optional object bstrPassword</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisProtection Protection(object bstrPassword)
		{
			return get_Protection(bstrPassword);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_HeaderMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "HeaderMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_HeaderMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "HeaderMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_HeaderMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double HeaderMargin(object unitsNameOrCode)
		{
			return get_HeaderMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_FooterMargin(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "FooterMargin", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
        /// <param name="value">optional double value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_FooterMargin(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "FooterMargin", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_FooterMargin
		/// </summary>
		/// <param name="unitsNameOrCode">optional object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double FooterMargin(object unitsNameOrCode)
		{
			return get_FooterMargin(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_Password(object bstrExistingPassword)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(bstrExistingPassword);
			object returnItem = Invoker.PropertyGet(this, "Password", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

        /// <summary>
        /// SupportByVersion Visio 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
        /// <param name="value">optional string value</param>
        [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Password(object bstrExistingPassword, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrExistingPassword);
			Invoker.PropertySet(this, "Password", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Password
		/// </summary>
		/// <param name="bstrExistingPassword">optional object bstrExistingPassword</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Password(object bstrExistingPassword)
		{
			return get_Password(bstrExistingPassword);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// Interface LPVISIODOCUMENT 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIODOCUMENT : COMObject
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
                    _type = typeof(LPVISIODOCUMENT);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIODOCUMENT(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIODOCUMENT(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 InPlace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InPlace", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMasters Masters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Masters", paramsArray);
				NetOffice.VisioApi.IVMasters newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMasters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPages Pages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Pages", paramsArray);
				NetOffice.VisioApi.IVPages newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVStyles Styles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Styles", paramsArray);
				NetOffice.VisioApi.IVStyles newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVStyles;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_Saved", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_Saved", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ReadOnly
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ReadOnly", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 old_Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_Version", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_Version", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
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
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Subject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Subject", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Subject", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Creator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Keywords
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Keywords", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Keywords", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Description
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Description", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Description", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVUIObject CustomMenus
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomMenus", paramsArray);
				NetOffice.VisioApi.IVUIObject newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVUIObject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string CustomMenusFile
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomMenusFile", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CustomMenusFile", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVUIObject CustomToolbars
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomToolbars", paramsArray);
				NetOffice.VisioApi.IVUIObject newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVUIObject;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string CustomToolbarsFile
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomToolbarsFile", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CustomToolbarsFile", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVFonts Fonts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Fonts", paramsArray);
				NetOffice.VisioApi.IVFonts newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVFonts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVColors Colors
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Colors", paramsArray);
				NetOffice.VisioApi.IVColors newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVColors;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EventList", paramsArray);
				NetOffice.VisioApi.IVEventList newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVEventList;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Template
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Template", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_SavePreviewMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_SavePreviewMode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_SavePreviewMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double LeftMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LeftMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LeftMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double RightMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RightMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RightMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double TopMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TopMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TopMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double BottomMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BottomMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BottomMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_PrintLandscape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_PrintLandscape", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_PrintLandscape", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_PrintCenteredH
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_PrintCenteredH", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_PrintCenteredH", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_PrintCenteredV
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_PrintCenteredV", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_PrintCenteredV", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double PrintScale
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintScale", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintScale", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_PrintFitOnPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_PrintFitOnPages", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_PrintFitOnPages", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 PrintPagesAcross
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintPagesAcross", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintPagesAcross", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 PrintPagesDown
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintPagesDown", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintPagesDown", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string DefaultStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string DefaultLineStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultLineStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultLineStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string DefaultFillStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultFillStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultFillStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string DefaultTextStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultTextStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultTextStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersistsEvents", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object VBProject
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "VBProject", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_PaperWidth(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "PaperWidth", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_PaperWidth
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double PaperWidth(object unitsNameOrCode)
		{
			return get_PaperWidth(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_PaperHeight(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "PaperHeight", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_PaperHeight
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double PaperHeight(object unitsNameOrCode)
		{
			return get_PaperHeight(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_PaperSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_PaperSize", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_PaperSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CodeName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CodeName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 old_Mode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_Mode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "old_Mode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVOLEObjects OLEObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OLEObjects", paramsArray);
				NetOffice.VisioApi.IVOLEObjects newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVOLEObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Manager
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Manager", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Manager", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Company
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Company", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Company", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Category
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Category", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Category", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string HyperlinkBase
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HyperlinkBase", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HyperlinkBase", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DocumentSheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DocumentSheet", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object Container
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Container", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string ClassID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClassID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
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
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMasterShortcuts MasterShortcuts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MasterShortcuts", paramsArray);
				NetOffice.VisioApi.IVMasterShortcuts newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMasterShortcuts;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string AlternateNames
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AlternateNames", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AlternateNames", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape GestureFormatSheet
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GestureFormatSheet", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool AutoRecover
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoRecover", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AutoRecover", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Saved", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Saved", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisDocVersions Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisDocVersions)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Version", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisSavePreviewMode SavePreviewMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SavePreviewMode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisSavePreviewMode)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SavePreviewMode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool PrintLandscape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintLandscape", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintLandscape", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool PrintCenteredH
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintCenteredH", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintCenteredH", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool PrintCenteredV
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintCenteredV", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintCenteredV", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool PrintFitOnPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintFitOnPages", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintFitOnPages", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisPaperSizes PaperSize
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PaperSize", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisPaperSizes)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PaperSize", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisDocModeArgs Mode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisDocModeArgs)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Mode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool SnapEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisSnapSettings SnapSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapSettings", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisSnapSettings)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapSettings", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisSnapExtensions SnapExtensions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SnapExtensions", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisSnapExtensions)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapExtensions", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double[] SnapAngles
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "SnapAngles", paramsArray);
				return (Double[])returnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "SnapAngles", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool GlueEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GlueEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GlueEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisGlueSettings GlueSettings
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GlueSettings", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisGlueSettings)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GlueSettings", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool DynamicGridEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DynamicGridEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DynamicGridEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string DefaultGuideStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultGuideStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultGuideStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisProtection Protection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Protection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisProtection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Protection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Printer
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Printer", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Printer", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 PrintCopies
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PrintCopies", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PrintCopies", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string HeaderLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderLeft", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string HeaderCenter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderCenter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderCenter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string HeaderRight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderRight", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderRight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double HeaderMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FooterLeft
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FooterLeft", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FooterLeft", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FooterCenter
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FooterCenter", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FooterCenter", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FooterRight
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FooterRight", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FooterRight", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double FooterMargin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FooterMargin", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FooterMargin", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public stdole.Font HeaderFooterFont
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderFooterFont", paramsArray);
				stdole.Font newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Font;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderFooterFont", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 HeaderFooterColor
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "HeaderFooterColor", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "HeaderFooterColor", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string Password
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Password", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Password", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public stdole.Picture PreviewPicture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreviewPicture", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PreviewPicture", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 BuildNumberCreated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildNumberCreated", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 BuildNumberEdited
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildNumberEdited", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public DateTime TimeCreated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimeCreated", paramsArray);
				return NetRuntimeSystem.Convert.ToDateTime(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public DateTime Time
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Time", paramsArray);
				return NetRuntimeSystem.Convert.ToDateTime(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public DateTime TimeEdited
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimeEdited", paramsArray);
				return NetRuntimeSystem.Convert.ToDateTime(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public DateTime TimePrinted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimePrinted", paramsArray);
				return NetRuntimeSystem.Convert.ToDateTime(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public DateTime TimeSaved
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TimeSaved", paramsArray);
				return NetRuntimeSystem.Convert.ToDateTime(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool ContainsWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainsWorkspace", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object[] EmailRoutingData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EmailRoutingData", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
                    ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this, (object[])returnItem);
					return newObject;
				}
				else
				{
					return (object[]) returnItem;
				}
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public byte[] VBProjectData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "VBProjectData", paramsArray);
				return (byte[])returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 SolutionXMLElementCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SolutionXMLElementCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_SolutionXMLElementName(Int32 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "SolutionXMLElementName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SolutionXMLElementName
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string SolutionXMLElementName(Int32 index)
		{
			return get_SolutionXMLElementName(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool get_SolutionXMLElementExists(string elementName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(elementName);
			object returnItem = Invoker.PropertyGet(this, "SolutionXMLElementExists", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SolutionXMLElementExists
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool SolutionXMLElementExists(string elementName)
		{
			return get_SolutionXMLElementExists(elementName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_SolutionXMLElement(string elementName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(elementName);
			object returnItem = Invoker.PropertyGet(this, "SolutionXMLElement", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_SolutionXMLElement(string elementName, string value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(elementName);
			Invoker.PropertySet(this, "SolutionXMLElement", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SolutionXMLElement
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string SolutionXMLElement(string elementName)
		{
			return get_SolutionXMLElement(elementName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 FullBuildNumberCreated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullBuildNumberCreated", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 FullBuildNumberEdited
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullBuildNumberEdited", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool MacrosEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MacrosEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisZoomBehavior ZoomBehavior
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ZoomBehavior", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisZoomBehavior)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ZoomBehavior", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.Enums.VisDocumentTypes Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.VisioApi.Enums.VisDocumentTypes)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 Language
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Language", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Language", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool UndoEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UndoEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UndoEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object SharedWorkspace
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SharedWorkspace", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object Sync
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Sync", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVDataRecordsets DataRecordsets
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataRecordsets", paramsArray);
				NetOffice.VisioApi.IVDataRecordsets newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDataRecordsets;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public bool ContainsWorkspaceEx
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainsWorkspaceEx", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ContainsWorkspaceEx", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public string DefaultSavePath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultSavePath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultSavePath", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public string CustomUI
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CustomUI", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CustomUI", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public string UserCustomUI
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UserCustomUI", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UserCustomUI", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVServerPublishOptions ServerPublishOptions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ServerPublishOptions", paramsArray);
				NetOffice.VisioApi.IVServerPublishOptions newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVServerPublishOptions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVValidation Validation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Validation", paramsArray);
				NetOffice.VisioApi.IVValidation newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVValidation;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32 DiagramServicesEnabled
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DiagramServicesEnabled", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DiagramServicesEnabled", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public bool CompatibilityMode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CompatibilityMode", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVComments Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				NetOffice.VisioApi.IVComments newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVComments;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="xPos">Int16 xPos</param>
		/// <param name="yPos">Int16 yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMaster Drop(object objectToDrop, Int16 xPos, Int16 yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, xPos, yPos);
			object returnItem = Invoker.MethodReturn(this, "Drop", paramsArray);
			NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Save()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Save", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SaveAs(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "SaveAs", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Print()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Print", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="menusObject">NetOffice.VisioApi.IVUIObject MenusObject</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetCustomMenus(NetOffice.VisioApi.IVUIObject menusObject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(menusObject);
			Invoker.Method(this, "SetCustomMenus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ClearCustomMenus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearCustomMenus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="toolbarsObject">NetOffice.VisioApi.IVUIObject ToolbarsObject</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetCustomToolbars(NetOffice.VisioApi.IVUIObject toolbarsObject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(toolbarsObject);
			Invoker.Method(this, "SetCustomToolbars", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ClearCustomToolbars()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearCustomToolbars", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="saveFlags">Int16 SaveFlags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SaveAsEx(string fileName, Int16 saveFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, saveFlags);
			Invoker.Method(this, "SaveAsEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="iD">Int16 ID</param>
		/// <param name="fileName">string FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetIcon(Int16 iD, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD, fileName);
			Invoker.Method(this, "GetIcon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="iD">Int16 ID</param>
		/// <param name="index">Int16 Index</param>
		/// <param name="fileName">string FileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetIcon(Int16 iD, Int16 index, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD, index, fileName);
			Invoker.Method(this, "SetIcon", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow OpenStencilWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OpenStencilWindow", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="line">string Line</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ParseLine(string line)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(line);
			Invoker.Method(this, "ParseLine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="line">string Line</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ExecuteLine(string line)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(line);
			Invoker.Method(this, "ExecuteLine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="target">string Target</param>
		/// <param name="location">string Location</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink45(string target, string location)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(target, location);
			Invoker.Method(this, "FollowHyperlink45", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="frame">optional object Frame</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="res1">optional object res1</param>
		/// <param name="res2">optional object res2</param>
		/// <param name="res3">optional object res3</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2, object res3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo, frame, newWindow, res1, res2, res3);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="frame">optional object Frame</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo, frame);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="frame">optional object Frame</param>
		/// <param name="newWindow">optional object NewWindow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo, frame, newWindow);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="frame">optional object Frame</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="res1">optional object res1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo, frame, newWindow, res1);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="address">string Address</param>
		/// <param name="subAddress">string SubAddress</param>
		/// <param name="extraInfo">optional object ExtraInfo</param>
		/// <param name="frame">optional object Frame</param>
		/// <param name="newWindow">optional object NewWindow</param>
		/// <param name="res1">optional object res1</param>
		/// <param name="res2">optional object res2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FollowHyperlink(string address, string subAddress, object extraInfo, object frame, object newWindow, object res1, object res2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(address, subAddress, extraInfo, frame, newWindow, res1, res2);
			Invoker.Method(this, "FollowHyperlink", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ClearGestureFormatSheet()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ClearGestureFormatSheet", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		/// <param name="nAlerts">optional object nAlerts</param>
		/// <param name="nFixes">optional object nFixes</param>
		/// <param name="bStopOnError">optional object bStopOnError</param>
		/// <param name="bLogFileName">optional object bLogFileName</param>
		/// <param name="nReserved">optional object nReserved</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName, object nReserved)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions, nAlerts, nFixes, bStopOnError, bLogFileName, nReserved);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		/// <param name="nAlerts">optional object nAlerts</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions, object nAlerts)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions, nAlerts);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		/// <param name="nAlerts">optional object nAlerts</param>
		/// <param name="nFixes">optional object nFixes</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions, object nAlerts, object nFixes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions, nAlerts, nFixes);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		/// <param name="nAlerts">optional object nAlerts</param>
		/// <param name="nFixes">optional object nFixes</param>
		/// <param name="bStopOnError">optional object bStopOnError</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions, nAlerts, nFixes, bStopOnError);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nTargets">optional object nTargets</param>
		/// <param name="nActions">optional object nActions</param>
		/// <param name="nAlerts">optional object nAlerts</param>
		/// <param name="nFixes">optional object nFixes</param>
		/// <param name="bStopOnError">optional object bStopOnError</param>
		/// <param name="bLogFileName">optional object bLogFileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Clean(object nTargets, object nActions, object nAlerts, object nFixes, object bStopOnError, object bLogFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nTargets, nActions, nAlerts, nFixes, bStopOnError, bLogFileName);
			Invoker.Method(this, "Clean", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pSourceDoc">NetOffice.VisioApi.IVDocument pSourceDoc</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CopyPreviewPicture(NetOffice.VisioApi.IVDocument pSourceDoc)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pSourceDoc);
			Invoker.Method(this, "CopyPreviewPicture", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="elementName">string ElementName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void DeleteSolutionXMLElement(string elementName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(elementName);
			Invoker.Method(this, "DeleteSolutionXMLElement", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool CanCheckIn()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanCheckIn", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		/// <param name="makePublic">optional bool MakePublic = false</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments, object makePublic)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments, makePublic);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CheckIn()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CheckIn(object saveChanges)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object Comments</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CheckIn(object saveChanges, object comments)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(saveChanges, comments);
			Invoker.Method(this, "CheckIn", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		/// <param name="printToFile">optional bool PrintToFile = false</param>
		/// <param name="outputFileName">optional string OutputFileName = </param>
		/// <param name="copies">optional Int32 Copies = 1</param>
		/// <param name="collate">optional bool Collate = false</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate, object colorAsBlack)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies, collate, colorAsBlack);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		/// <param name="printToFile">optional bool PrintToFile = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		/// <param name="printToFile">optional bool PrintToFile = false</param>
		/// <param name="outputFileName">optional string OutputFileName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		/// <param name="printToFile">optional bool PrintToFile = false</param>
		/// <param name="outputFileName">optional string OutputFileName = </param>
		/// <param name="copies">optional Int32 Copies = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="scaleCurrentViewToPaper">optional bool ScaleCurrentViewToPaper = false</param>
		/// <param name="printerName">optional string PrinterName = </param>
		/// <param name="printToFile">optional bool PrintToFile = false</param>
		/// <param name="outputFileName">optional string OutputFileName = </param>
		/// <param name="copies">optional Int32 Copies = 1</param>
		/// <param name="collate">optional bool Collate = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PrintOut(NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object scaleCurrentViewToPaper, object printerName, object printToFile, object outputFileName, object copies, object collate)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(printRange, fromPage, toPage, scaleCurrentViewToPaper, printerName, printToFile, outputFileName, copies, collate);
			Invoker.Method(this, "PrintOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrUndoScopeName">string bstrUndoScopeName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 BeginUndoScope(string bstrUndoScopeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrUndoScopeName);
			object returnItem = Invoker.MethodReturn(this, "BeginUndoScope", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="nScopeID">Int32 nScopeID</param>
		/// <param name="bCommit">bool bCommit</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void EndUndoScope(Int32 nScopeID, bool bCommit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(nScopeID, bCommit);
			Invoker.Method(this, "EndUndoScope", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="pUndoUnit">object pUndoUnit</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddUndoUnit(object pUndoUnit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pUndoUnit);
			Invoker.Method(this, "AddUndoUnit", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PurgeUndo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PurgeUndo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="bstrScopeName">string bstrScopeName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void RenameCurrentScope(string bstrScopeName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrScopeName);
			Invoker.Method(this, "RenameCurrentScope", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="removeHiddenInfoItems">Int32 RemoveHiddenInfoItems</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void RemoveHiddenInformation(Int32 removeHiddenInfoItems)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(removeHiddenInfoItems);
			Invoker.Method(this, "RemoveHiddenInformation", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
		/// <param name="nameArray">String[] NameArray</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetThemeNames(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(eType, (object)nameArray);
			Invoker.Method(this, "GetThemeNames", paramsArray, modifiers);
			nameArray = (String[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="eType">NetOffice.VisioApi.Enums.VisThemeTypes eType</param>
		/// <param name="nameArray">String[] NameArray</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetThemeNamesU(NetOffice.VisioApi.Enums.VisThemeTypes eType, out String[] nameArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			nameArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray(eType, (object)nameArray);
			Invoker.Method(this, "GetThemeNamesU", paramsArray, modifiers);
			nameArray = (String[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public bool CanUndoCheckOut()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CanUndoCheckOut", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void UndoCheckOut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UndoCheckOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		/// <param name="includeBackground">optional bool IncludeBackground = true</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="fixedFormatExtClass">optional object FixedFormatExtClass</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1, object fixedFormatExtClass)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags, useISO19005_1, fixedFormatExtClass);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		/// <param name="includeBackground">optional bool IncludeBackground = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		/// <param name="includeBackground">optional bool IncludeBackground = true</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		/// <param name="includeBackground">optional bool IncludeBackground = true</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fixedFormat">NetOffice.VisioApi.Enums.VisFixedFormatTypes FixedFormat</param>
		/// <param name="outputFileName">string OutputFileName</param>
		/// <param name="intent">NetOffice.VisioApi.Enums.VisDocExIntent Intent</param>
		/// <param name="printRange">NetOffice.VisioApi.Enums.VisPrintOutRange PrintRange</param>
		/// <param name="fromPage">optional Int32 FromPage = 1</param>
		/// <param name="toPage">optional Int32 ToPage = -1</param>
		/// <param name="colorAsBlack">optional bool ColorAsBlack = false</param>
		/// <param name="includeBackground">optional bool IncludeBackground = true</param>
		/// <param name="includeDocumentProperties">optional bool IncludeDocumentProperties = true</param>
		/// <param name="includeStructureTags">optional bool IncludeStructureTags = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void ExportAsFixedFormat(NetOffice.VisioApi.Enums.VisFixedFormatTypes fixedFormat, string outputFileName, NetOffice.VisioApi.Enums.VisDocExIntent intent, NetOffice.VisioApi.Enums.VisPrintOutRange printRange, object fromPage, object toPage, object colorAsBlack, object includeBackground, object includeDocumentProperties, object includeStructureTags, object useISO19005_1)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fixedFormat, outputFileName, intent, printRange, fromPage, toPage, colorAsBlack, includeBackground, includeDocumentProperties, includeStructureTags, useISO19005_1);
			Invoker.Method(this, "ExportAsFixedFormat", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}