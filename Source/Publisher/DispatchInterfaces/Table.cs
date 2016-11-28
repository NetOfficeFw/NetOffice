using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.PublisherApi
{
	///<summary>
	/// Table
	///</summary>
	public class Table_ : COMObject
	{
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		/// <param name="endRow">optional Int32 EndRow</param>
		/// <param name="endColumn">optional Int32 EndColumn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow, object endColumn)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startRow, startColumn, endRow, endColumn);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.PublisherApi.CellRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.CellRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		/// <param name="endRow">optional Int32 EndRow</param>
		/// <param name="endColumn">optional Int32 EndColumn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow, object endColumn)
		{
			return get_Cells(startRow, startColumn, endRow, endColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startRow);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.PublisherApi.CellRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.CellRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells(object startRow)
		{
			return get_Cells(startRow);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startRow, startColumn);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.PublisherApi.CellRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.CellRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn)
		{
			return get_Cells(startRow, startColumn);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		/// <param name="endRow">optional Int32 EndRow</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.PublisherApi.CellRange get_Cells(object startRow, object startColumn, object endRow)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(startRow, startColumn, endRow);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.PublisherApi.CellRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.CellRange;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="startRow">optional Int32 StartRow</param>
		/// <param name="startColumn">optional Int32 StartColumn</param>
		/// <param name="endRow">optional Int32 EndRow</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells(object startRow, object startColumn, object endRow)
		{
			return get_Cells(startRow, startColumn, endRow);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface Table 
	/// SupportByVersion Publisher, 14,15,16
	///</summary>
	[SupportByVersionAttribute("Publisher", 14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Table : Table_
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
                    _type = typeof(Table);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.PublisherApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Application.LateBindingApiWrapperType) as NetOffice.PublisherApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Columns Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.PublisherApi.Columns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Columns.LateBindingApiWrapperType) as NetOffice.PublisherApi.Columns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public bool GrowToFitText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GrowToFitText", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GrowToFitText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Rows Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.PublisherApi.Rows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.Rows.LateBindingApiWrapperType) as NetOffice.PublisherApi.Rows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.PbTableDirectionType TableDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.PublisherApi.Enums.PbTableDirectionType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TableDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.CellRange Cells
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
				NetOffice.PublisherApi.CellRange newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.PublisherApi.CellRange.LateBindingApiWrapperType) as NetOffice.PublisherApi.CellRange;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType AutoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		/// <param name="fill">optional bool Fill = true</param>
		/// <param name="borders">optional bool Borders = true</param>
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill, object borders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(autoFormat, textFormatting, textAlignment, fill, borders);
			Invoker.Method(this, "ApplyAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType AutoFormat</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(autoFormat);
			Invoker.Method(this, "ApplyAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType AutoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(autoFormat, textFormatting);
			Invoker.Method(this, "ApplyAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType AutoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(autoFormat, textFormatting, textAlignment);
			Invoker.Method(this, "ApplyAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// 
		/// </summary>
		/// <param name="autoFormat">NetOffice.PublisherApi.Enums.PbTableAutoFormatType AutoFormat</param>
		/// <param name="textFormatting">optional bool TextFormatting = true</param>
		/// <param name="textAlignment">optional bool TextAlignment = true</param>
		/// <param name="fill">optional bool Fill = true</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Publisher", 14,15,16)]
		public void ApplyAutoFormat(NetOffice.PublisherApi.Enums.PbTableAutoFormatType autoFormat, object textFormatting, object textAlignment, object fill)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(autoFormat, textFormatting, textAlignment, fill);
			Invoker.Method(this, "ApplyAutoFormat", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}