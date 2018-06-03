using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface TableStyle 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194537.aspx </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TableStyle : COMObject, NetOffice.ExcelApi.TableStyle
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
                    _type = typeof(TableStyle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TableStyle() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821286.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823073.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820983.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string _Default
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198227.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839881.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public string NameLocal
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameLocal");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840826.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool BuiltIn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BuiltIn");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192949.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.TableStyleElements TableStyleElements
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TableStyleElements>(this, "TableStyleElements", typeof(NetOffice.ExcelApi.TableStyleElements));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839252.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowAsAvailableTableStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAsAvailableTableStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAsAvailableTableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839032.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ShowAsAvailablePivotTableStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAsAvailablePivotTableStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAsAvailablePivotTableStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839730.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ShowAsAvailableSlicerStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAsAvailableSlicerStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAsAvailableSlicerStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231088.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool ShowAsAvailableTimelineStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAsAvailableTimelineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAsAvailableTimelineStyle", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821506.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821658.aspx </remarks>
		/// <param name="newTableStyleName">optional object newTableStyleName</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.TableStyle Duplicate(object newTableStyleName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.TableStyle>(this, "Duplicate", typeof(NetOffice.ExcelApi.TableStyle), newTableStyleName);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821658.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.TableStyle Duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.TableStyle>(this, "Duplicate", typeof(NetOffice.ExcelApi.TableStyle));
		}

		#endregion

		#pragma warning restore
	}
}


