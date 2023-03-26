using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _Report 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method)]
	public class _Report : COMObject, IEnumerableProvider<object>
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
                    _type = typeof(_Report);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Report(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Report(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Report(string progId) : base(progId)
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
		public string FormName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.RecordSource"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string RecordSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RecordSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Filter(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Filter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FilterOn"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FilterOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FilterOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FilterOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OrderBy"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OrderBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OrderBy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OrderByOn"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool OrderByOn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OrderByOn");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrderByOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ServerFilter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ServerFilter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ServerFilter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ServerFilter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Caption"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.RecordLocks"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte RecordLocks
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "RecordLocks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecordLocks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PageHeader"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PageHeader
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PageHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PageFooter"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PageFooter
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PageFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DateGrouping"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte DateGrouping
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "DateGrouping");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DateGrouping", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.GrpKeepTogether"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte GrpKeepTogether
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "GrpKeepTogether");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GrpKeepTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MinButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MinButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MinButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MaxButton
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MaxButton");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxButton", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Width"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Width
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Width");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Width", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Picture"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Picture
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Picture");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Picture", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PictureType"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureType
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PictureSizeMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureSizeMode
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureSizeMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureSizeMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PictureAlignment"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PictureAlignment
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PictureAlignment");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PictureTiling"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PictureTiling
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PictureTiling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureTiling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PicturePages"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte PicturePages
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "PicturePages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PicturePages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.MenuBar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string MenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Toolbar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Toolbar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Toolbar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Toolbar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ShortcutMenuBar"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string ShortcutMenuBar
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.GridX"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridX
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.GridY"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 GridY
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.LayoutForPrint"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool LayoutForPrint
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LayoutForPrint");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LayoutForPrint", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FastLaserPrinting"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool FastLaserPrinting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FastLaserPrinting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FastLaserPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.HelpFile"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string HelpFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HelpFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.HelpContextId"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 HelpContextId
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HelpContextId");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HelpContextId", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Hwnd"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Hwnd
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hwnd");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hwnd", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Count"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Count
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Count");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Count", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Page(property)"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Page
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Page");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Page", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Pages"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 Pages
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Pages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Pages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageWidth
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LogicalPageHeight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LogicalPageHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LogicalPageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ZoomControl
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ZoomControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ZoomControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.HasData"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 HasData
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HasData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Left"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Left
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Left");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Left", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Top"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Top
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Top");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Top", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Height"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Height
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Height");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Height", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PrintSection"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool PrintSection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintSection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintSection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.NextRecord"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool NextRecord
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NextRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NextRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.MoveLayout"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool MoveLayout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MoveLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MoveLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FormatCount"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FormatCount
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FormatCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormatCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PrintCount"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 PrintCount
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PrintCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Visible"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Painting"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Painting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Painting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Painting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PrtMip"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtMip
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtMip");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtMip", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PrtDevMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevMode
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevMode");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PrtDevNames"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PrtDevNames
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrtDevNames");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrtDevNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ForeColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.CurrentX"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single CurrentX
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CurrentX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.CurrentY"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single CurrentY
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CurrentY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ScaleHeight"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single ScaleHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScaleHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ScaleLeft"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single ScaleLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScaleLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ScaleMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 ScaleMode
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ScaleMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ScaleTop"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single ScaleTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScaleTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ScaleWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single ScaleWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ScaleWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FontBold"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FontBold
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FontItalic"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FontItalic
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FontName"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FontSize"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FontSize
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FontUnderline"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FontUnderline
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DrawMode"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DrawMode
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DrawMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DrawStyle"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DrawStyle
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DrawStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DrawWidth"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 DrawWidth
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "DrawWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DrawWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FillColor"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 FillColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FillColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.FillStyle"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int16 FillStyle
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FillStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "_Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PaletteSource"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string PaletteSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PaletteSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PaletteSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Tag"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Tag
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Tag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PaintPalette"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PaintPalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PaintPalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PaintPalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string OnMenu
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnMenu");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnMenu", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnOpen"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnOpen
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnOpen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnOpen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnClose"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnClose
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnClose");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnClose", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnActivate"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnActivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnActivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnActivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnDeactivate"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnDeactivate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnDeactivate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnDeactivate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnNoData"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnNoData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnNoData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnNoData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnPage"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnPage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.OnError"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string OnError
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OnError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OnError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Dirty"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool Dirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Dirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.CurrentRecord"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 CurrentRecord
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurrentRecord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurrentRecord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PictureData"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PictureData
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PictureData");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PictureData", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PicturePalette"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PicturePalette
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PicturePalette");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PicturePalette", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.HasModule"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool HasModule
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasModule");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HasModule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 acHiddenCurrentPage
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "acHiddenCurrentPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "acHiddenCurrentPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Orientation"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public byte Orientation
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.InputParameters"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string InputParameters
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InputParameters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InputParameters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Application"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Parent"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.ActiveControl"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Control ActiveControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "ActiveControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DefaultControl"/> </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Control get_DefaultControl(Int32 controlType)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Control>(this, "DefaultControl", NetOffice.AccessApi.Control.LateBindingApiWrapperType, controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DefaultControl
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.DefaultControl"/> </remarks>
		/// <param name="controlType">Int32 controlType</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_DefaultControl")]
		public NetOffice.AccessApi.Control DefaultControl(Int32 controlType)
		{
			return get_DefaultControl(controlType);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Section"/> </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.Section get_Section(object index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Section>(this, "Section", NetOffice.AccessApi.Section.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Section"/> </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_Section")]
		public NetOffice.AccessApi.Section Section(object index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.GroupLevel"/> </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.AccessApi.GroupLevel get_GroupLevel(Int32 index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.GroupLevel>(this, "GroupLevel", NetOffice.AccessApi.GroupLevel.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_GroupLevel
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.GroupLevel"/> </remarks>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), Redirect("get_GroupLevel")]
		public NetOffice.AccessApi.GroupLevel GroupLevel(Int32 index)
		{
			return get_GroupLevel(index);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Report"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Report Report
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Report>(this, "Report", NetOffice.AccessApi.Report.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Module"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Module Module
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Module>(this, "Module", NetOffice.AccessApi.Module.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Properties"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", NetOffice.AccessApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Controls"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Controls Controls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Controls>(this, "Controls", NetOffice.AccessApi.Controls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Name"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Circle"/> </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		/// <param name="radius">Single radius</param>
		/// <param name="color">Int32 color</param>
		/// <param name="start">Single start</param>
		/// <param name="end">Single end</param>
		/// <param name="aspect">Single aspect</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Circle(Int16 flags, Single x, Single y, Single radius, Int32 color, Single start, Single end, Single aspect)
		{
			 Factory.ExecuteMethod(this, "Circle", new object[]{ flags, x, y, radius, color, start, end, aspect });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Line"/> </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Line(Int16 flags, Single x1, Single y1, Single x2, Single y2, Int32 color)
		{
			 Factory.ExecuteMethod(this, "Line", new object[]{ flags, x1, y1, x2, y2, color });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.PSet"/> </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x">Single x</param>
		/// <param name="y">Single y</param>
		/// <param name="color">Int32 color</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void PSet(Int16 flags, Single x, Single y, Int32 color)
		{
			 Factory.ExecuteMethod(this, "PSet", flags, x, y, color);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Scale"/> </remarks>
		/// <param name="flags">Int16 flags</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		/// <param name="x2">Single x2</param>
		/// <param name="y2">Single y2</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Scale(Int16 flags, Single x1, Single y1, Single x2, Single y2)
		{
			 Factory.ExecuteMethod(this, "Scale", new object[]{ flags, x1, y1, x2, y2 });
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.TextWidth"/> </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single TextWidth(string expr)
		{
			return Factory.ExecuteSingleMethodGet(this, "TextWidth", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.TextHeight"/> </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Single TextHeight(string expr)
		{
			return Factory.ExecuteSingleMethodGet(this, "TextHeight", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.Report.Print"/> </remarks>
		/// <param name="expr">string expr</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void Print(string expr)
		{
			 Factory.ExecuteMethod(this, "Print", expr);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="bstrExpr">string bstrExpr</param>
		/// <param name="ppsa">optional object[] ppsa</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public object _Evaluate(string bstrExpr, object[] ppsa)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
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
		public object _Evaluate(string bstrExpr)
		{
			return Factory.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
		}

        #endregion
       
        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
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
        public IEnumerator<object> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}
