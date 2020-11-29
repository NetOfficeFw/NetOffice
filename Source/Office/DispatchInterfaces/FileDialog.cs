using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface FileDialog 
	/// SupportByVersion Office, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog"/> </remarks>
	[SupportByVersion("Office", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class FileDialog : _IMsoDispObj
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
                    _type = typeof(FileDialog);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FileDialog(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FileDialog(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FileDialog(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Parent"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Filters"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileDialogFilters Filters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialogFilters>(this, "Filters", NetOffice.OfficeApi.FileDialogFilters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.FilterIndex"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 FilterIndex
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FilterIndex");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Title"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
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
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.ButtonName"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string ButtonName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ButtonName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ButtonName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.AllowMultiSelect"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public bool AllowMultiSelect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowMultiSelect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowMultiSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.InitialView"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileDialogView InitialView
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileDialogView>(this, "InitialView");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "InitialView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.InitialFileName"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string InitialFileName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InitialFileName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InitialFileName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.SelectedItems"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FileDialogSelectedItems SelectedItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialogSelectedItems>(this, "SelectedItems", NetOffice.OfficeApi.FileDialogSelectedItems.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.DialogType"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoFileDialogType DialogType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileDialogType>(this, "DialogType");
			}
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Item"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public string Item
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Item");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Show"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public Int32 Show()
		{
			return Factory.ExecuteInt32MethodGet(this, "Show");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.FileDialog.Execute"/> </remarks>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public void Execute()
		{
			 Factory.ExecuteMethod(this, "Execute");
		}

		#endregion

		#pragma warning restore
	}
}
