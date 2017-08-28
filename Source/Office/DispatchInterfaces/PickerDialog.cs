using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface PickerDialog 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860858.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PickerDialog : _IMsoDispObj
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
                    _type = typeof(PickerDialog);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PickerDialog(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PickerDialog(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862371.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public string DataHandlerId
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataHandlerId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataHandlerId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862526.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860248.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerProperties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerProperties>(this, "Properties", NetOffice.OfficeApi.PickerProperties.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861181.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults CreatePickerResults()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "CreatePickerResults", NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
		/// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
		/// <param name="existingResults">optional NetOffice.OfficeApi.PickerResults ExistingResults = 0</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show(object isMultiSelect, object existingResults)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType, isMultiSelect, existingResults);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861095.aspx </remarks>
		/// <param name="isMultiSelect">optional bool IsMultiSelect = true</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Show(object isMultiSelect)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Show", NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType, isMultiSelect);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861733.aspx </remarks>
		/// <param name="tokenText">string tokenText</param>
		/// <param name="duplicateDlgMode">Int32 duplicateDlgMode</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PickerResults Resolve(string tokenText, Int32 duplicateDlgMode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PickerResults>(this, "Resolve", NetOffice.OfficeApi.PickerResults.LateBindingApiWrapperType, tokenText, duplicateDlgMode);
		}

		#endregion

		#pragma warning restore
	}
}
