using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _Report 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>	
	public class _Report : COMObject, NetOffice.AccessApi._Report
    {
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi._Report);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(_Report);                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Report() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string FormName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821454.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string RecordSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834497.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197677.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool FilterOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FilterOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845637.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OrderBy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OrderBy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OrderBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197033.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool OrderByOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OrderByOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OrderByOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836003.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ServerFilter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197377.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Caption
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191707.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte RecordLocks
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "RecordLocks");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RecordLocks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820727.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PageHeader
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PageHeader");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196711.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PageFooter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PageFooter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835694.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte DateGrouping
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DateGrouping");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DateGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194805.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte GrpKeepTogether
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GrpKeepTogether");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GrpKeepTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool MinButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MinButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MinButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual bool MaxButton
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MaxButton");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836874.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Width
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Width");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845621.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Picture
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Picture");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Picture", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197756.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195853.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureSizeMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureSizeMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834708.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PictureAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193158.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool PictureTiling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820802.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte PicturePages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PicturePages");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PicturePages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844730.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string MenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836048.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Toolbar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Toolbar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Toolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191912.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string ShortcutMenuBar
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822460.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 GridX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "GridX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835940.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 GridY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "GridY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836872.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool LayoutForPrint
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LayoutForPrint");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LayoutForPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822465.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool FastLaserPrinting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FastLaserPrinting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FastLaserPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834686.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string HelpFile
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192524.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 HelpContextId
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835728.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Hwnd
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hwnd");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hwnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195823.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195525.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Page
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Page");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822467.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 Pages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Pages");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Pages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 LogicalPageWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LogicalPageWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LogicalPageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 LogicalPageHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LogicalPageHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LogicalPageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 ZoomControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ZoomControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ZoomControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836046.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 HasData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HasData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195237.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Left
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822522.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Top
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845460.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 Height
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195874.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool PrintSection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintSection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintSection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196049.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool NextRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NextRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NextRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821781.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool MoveLayout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MoveLayout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MoveLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192467.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FormatCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FormatCount");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormatCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197646.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 PrintCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PrintCount");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822441.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Visible
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196705.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Painting
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Painting");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Painting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836632.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtMip
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtMip");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtMip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820784.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtDevMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtDevMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtDevMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197750.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PrtDevNames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PrtDevNames");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PrtDevNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844842.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 ForeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192663.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single CurrentX
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "CurrentX");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844825.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single CurrentY
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "CurrentY");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822004.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single ScaleHeight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleHeight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845811.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single ScaleLeft
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleLeft");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835754.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 ScaleMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ScaleMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192246.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single ScaleTop
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleTop");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822411.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single ScaleWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ScaleWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScaleWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845097.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FontBold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FontBold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835983.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FontItalic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FontItalic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192490.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string FontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196467.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FontSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FontSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192503.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FontUnderline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FontUnderline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196055.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 DrawMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DrawMode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845223.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 DrawStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DrawStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845724.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 DrawWidth
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "DrawWidth");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DrawWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844859.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 FillColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FillColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845296.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int16 FillStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FillStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string _Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198251.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string PaletteSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PaletteSource");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PaletteSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196443.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Tag
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835017.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PaintPalette
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PaintPalette");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PaintPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string OnMenu
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMenu");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835753.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnOpen
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnOpen");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194970.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnClose
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClose");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836259.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnActivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnActivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192078.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnDeactivate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDeactivate");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194603.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnNoData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnNoData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnNoData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835059.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191992.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string OnError
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnError");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835112.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool Dirty
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835318.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Int32 CurrentRecord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentRecord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822457.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PictureData
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PictureData");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PictureData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194186.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object PicturePalette
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PicturePalette");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "PicturePalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821121.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool HasModule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasModule");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasModule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 acHiddenCurrentPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "acHiddenCurrentPage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "acHiddenCurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193189.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual byte Orientation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "Orientation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823100.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string InputParameters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "InputParameters");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InputParameters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192257.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197107.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195791.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Control ActiveControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ActiveControl", typeof(NetOffice.AccessApi.Control));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845438.aspx </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Control get_DefaultControl(Int32 controlType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "DefaultControl", typeof(NetOffice.AccessApi.Control), controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultControl
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845438.aspx </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_DefaultControl")]
		public virtual NetOffice.AccessApi.Control DefaultControl(Int32 controlType)
		{
			return get_DefaultControl(controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192668.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.Section get_Section(object index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "Section", typeof(NetOffice.AccessApi.Section), index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192668.aspx </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public virtual NetOffice.AccessApi.Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197094.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.AccessApi.GroupLevel get_GroupLevel(Int32 index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.GroupLevel>(this, "GroupLevel", typeof(NetOffice.AccessApi.GroupLevel), index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_GroupLevel
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197094.aspx </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_GroupLevel")]
		public virtual NetOffice.AccessApi.GroupLevel GroupLevel(Int32 index)
		{
			return get_GroupLevel(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845176.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Report Report
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Report>(this, "Report", typeof(NetOffice.AccessApi.Report));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835656.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Module Module
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Module>(this, "Module", typeof(NetOffice.AccessApi.Module));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844949.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", typeof(NetOffice.AccessApi.Properties));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836083.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Controls Controls
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Controls>(this, "Controls", typeof(NetOffice.AccessApi.Controls));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821153.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193783.aspx </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		/// <param name="radius">Single radius</param>
		/// <param name="color">Int32 color</param>
		/// <param name="start">Single start</param>
		/// <param name="end">Single end</param>
		/// <param name="aspect">Single aspect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Circle(Int16 flags, Single x, Single y, Single radius, Int32 color, Single start, Single end, Single aspect)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Circle", new object[]{ flags, x, y, radius, color, start, end, aspect });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198297.aspx </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Line(Int16 flags, Single x1, Single y1, Single x2, Single y2, Int32 color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Line", new object[]{ flags, x1, y1, x2, y2, color });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197683.aspx </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void PSet(Int16 flags, Single x, Single y, Int32 color)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PSet", flags, x, y, color);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195398.aspx </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Scale(Int16 flags, Single x1, Single y1, Single x2, Single y2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Scale", new object[]{ flags, x1, y1, x2, y2 });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197941.aspx </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single TextWidth(string expr)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "TextWidth", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834360.aspx </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual Single TextHeight(string expr)
		{
			return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "TextHeight", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195586.aspx </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void Print(string expr)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Print", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object _Evaluate(string bstrExpr, object[] ppsa)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual object _Evaluate(string bstrExpr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

        #endregion
       
        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<object>

        /// <summary>
        /// SupportByVersion Access, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Access, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Access", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, true);
		}

		#endregion

		#pragma warning restore
	}
}


