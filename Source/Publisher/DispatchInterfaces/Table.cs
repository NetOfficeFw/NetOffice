using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// Table
	/// </summary>
	[SyntaxBypass]
 	public class Table_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Table_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Table_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		/// <param name="endRow">optional Int32 endRow</param>
		/// <param name="endColumn">optional Int32 endColumn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow, object endColumn)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType, startRow, startColumn, endRow, endColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		/// <param name="endRow">optional Int32 endRow</param>
		/// <param name="endColumn">optional Int32 endColumn</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_Cells")]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow, object endColumn)
		{
			return get_Cells(startRow, startColumn, endRow, endColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType, startRow);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_Cells")]
		public NetOffice.PublisherApi.CellRange Cells(object startRow)
		{
			return get_Cells(startRow);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType, startRow, startColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_Cells")]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn)
		{
			return get_Cells(startRow, startColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		/// <param name="endRow">optional Int32 endRow</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType, startRow, startColumn, endRow);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 startRow</param>
		/// <param name="startColumn">optional Int32 startColumn</param>
		/// <param name="endRow">optional Int32 endRow</param>
		[SupportByVersion("Publisher", 14,15,16), Redirect("get_Cells")]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow)
		{
			return get_Cells(startRow, startColumn, endRow);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface Table 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Table : Table_
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
                    _type = typeof(Table);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Table(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Table(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Columns Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Columns>(this, "Columns", NetOffice.PublisherApi.Columns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public bool GrowToFitText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GrowToFitText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GrowToFitText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Rows Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Rows>(this, "Rows", NetOffice.PublisherApi.Rows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbTableDirectionType TableDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbTableDirectionType>(this, "TableDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TableDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.CellRange>(this, "Cells", NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		/// <param name="fill">optional bool Fill = true</param>
		/// <param name="borders">optional bool Borders = true</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill, object borders)
		{
			 Factory.ExecuteMethod(this, "ApplyAutoFormat", new object[]{ autoFormat, textFormatting, textAlignment, fill, borders });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat)
		{
			 Factory.ExecuteMethod(this, "ApplyAutoFormat", autoFormat);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting)
		{
			 Factory.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment)
		{
			 Factory.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting, textAlignment);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		/// <param name="fill">optional bool Fill = true</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill)
		{
			 Factory.ExecuteMethod(this, "ApplyAutoFormat", autoFormat, textFormatting, textAlignment, fill);
		}

		#endregion

		#pragma warning restore
	}
}
