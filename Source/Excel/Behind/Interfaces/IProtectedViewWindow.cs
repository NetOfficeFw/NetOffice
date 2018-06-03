using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// Interface IProtectedViewWindow 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>
	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IProtectedViewWindow : COMObject, NetOffice.ExcelApi.IProtectedViewWindow
	{
		#pragma warning disable

		#region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IProtectedViewWindow);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IProtectedViewWindow() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string Caption
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool EnableResize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableResize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableResize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Height
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Left
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Top
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Double Width
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool Visible
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string SourceName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceName");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public string SourcePath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourcePath");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlProtectedViewWindowState WindowState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlProtectedViewWindowState>(this, "WindowState");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WindowState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Workbook Workbook
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbook>(this, "Workbook", typeof(NetOffice.ExcelApi.Workbook));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public Int32 Activate()
		{
			return Factory.ExecuteInt32MethodGet(this, "Activate");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public bool Close()
		{
			return Factory.ExecuteBoolMethodGet(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Workbook Edit(object writeResPassword, object updateLinks)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Edit", typeof(NetOffice.ExcelApi.Workbook), writeResPassword, updateLinks);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Workbook Edit()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Edit", typeof(NetOffice.ExcelApi.Workbook));
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Workbook Edit(object writeResPassword)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Edit", typeof(NetOffice.ExcelApi.Workbook), writeResPassword);
		}

		#endregion

		#pragma warning restore
	}
}


