using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OfficeApi
{
	///<summary>
	/// DispatchInterface PickerDialog 
	/// SupportByVersion Office, 14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860858.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PickerDialog : _IMsoDispObj
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
                    _type = typeof(PickerDialog);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PickerDialog(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862371.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public string DataHandlerId
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataHandlerId", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataHandlerId", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862526.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860248.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerProperties Properties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Properties", paramsArray);
				NetOffice.OfficeApi.PickerProperties newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OfficeApi.PickerProperties.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerProperties;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861181.aspx
		/// </summary>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults CreatePickerResults()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreatePickerResults", paramsArray);
			NetOffice.OfficeApi.PickerResults newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerResults;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx
		/// </summary>
		/// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
		/// <param name="existingResults">optional NetOffice.OfficeApi.PickerResults ExistingResults = 0</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show(object isMultiSelect, object existingResults)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(isMultiSelect, existingResults);
			object returnItem = Invoker.MethodReturn(this, "Show", paramsArray);
			NetOffice.OfficeApi.PickerResults newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerResults;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Show", paramsArray);
			NetOffice.OfficeApi.PickerResults newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerResults;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx
		/// </summary>
		/// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show(object isMultiSelect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(isMultiSelect);
			object returnItem = Invoker.MethodReturn(this, "Show", paramsArray);
			NetOffice.OfficeApi.PickerResults newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerResults;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861733.aspx
		/// </summary>
		/// <param name="tokenText">string TokenText</param>
		/// <param name="duplicateDlgMode">Int32 duplicateDlgMode</param>
		[SupportByVersionAttribute("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Resolve(string tokenText, Int32 duplicateDlgMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tokenText, duplicateDlgMode);
			object returnItem = Invoker.MethodReturn(this, "Resolve", paramsArray);
			NetOffice.OfficeApi.PickerResults newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType) as NetOffice.OfficeApi.PickerResults;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}