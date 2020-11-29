﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Protection 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Protection : COMObject
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
                    _type = typeof(Protection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Protection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Protection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Protection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowFormattingCells"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowFormattingCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFormattingCells");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowFormattingColumns"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowFormattingColumns
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFormattingColumns");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowFormattingRows"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowFormattingRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFormattingRows");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowInsertingColumns"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowInsertingColumns
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowInsertingColumns");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowInsertingRows"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowInsertingRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowInsertingRows");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowInsertingHyperlinks"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowInsertingHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowInsertingHyperlinks");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowDeletingColumns"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowDeletingColumns
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletingColumns");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowDeletingRows"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowDeletingRows
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowDeletingRows");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowSorting"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowSorting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowSorting");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowFiltering"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowFiltering
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFiltering");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowUsingPivotTables"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool AllowUsingPivotTables
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowUsingPivotTables");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Protection.AllowEditRanges"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.AllowEditRanges AllowEditRanges
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AllowEditRanges>(this, "AllowEditRanges", NetOffice.ExcelApi.AllowEditRanges.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
