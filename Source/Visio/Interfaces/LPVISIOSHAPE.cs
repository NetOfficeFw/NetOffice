using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// LPVISIOSHAPE
	///</summary>
	public class LPVISIOSHAPE_ : COMObject
	{
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOSHAPE_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_AreaIU(object fIncludeSubShapes)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fIncludeSubShapes);
			object returnItem = Invoker.PropertyGet(this, "AreaIU", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_AreaIU
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double AreaIU(object fIncludeSubShapes)
		{
			return get_AreaIU(fIncludeSubShapes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_LengthIU(object fIncludeSubShapes)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fIncludeSubShapes);
			object returnItem = Invoker.PropertyGet(this, "LengthIU", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_LengthIU
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double LengthIU(object fIncludeSubShapes)
		{
			return get_LengthIU(fIncludeSubShapes);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// Interface LPVISIOSHAPE 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIOSHAPE : COMObject
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
                    _type = typeof(LPVISIOSHAPE);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOSHAPE(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOSHAPE(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Document", paramsArray);
				NetOffice.VisioApi.IVDocument newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVDocument;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Parent", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.VisioApi.IVApplication newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVApplication;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Stat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMaster Master
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Master", paramsArray);
				NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeSpecificCellName);
			object returnItem = Invoker.PropertyGet(this, "Cells", paramsArray);
			NetOffice.VisioApi.IVCell newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName)
		{
			return get_Cells(localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="column">Int16 Column</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, row, column);
			object returnItem = Invoker.PropertyGet(this, "CellsSRC", paramsArray);
			NetOffice.VisioApi.IVCell newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRC
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="column">Int16 Column</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column)
		{
			return get_CellsSRC(section, row, column);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShapes Shapes
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shapes", paramsArray);
				NetOffice.VisioApi.IVShapes newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShapes;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Data1
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Data1", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Data1", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Data2
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Data2", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Data2", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Data3
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Data3", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Data3", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Help
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Help", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Help", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string NameID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NameID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Text", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 CharCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CharCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCharacters Characters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Characters", paramsArray);
				NetOffice.VisioApi.IVCharacters newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCharacters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 OneD
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OneD", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OneD", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 GeometryCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GeometryCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowCount(Int16 section)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section);
			object returnItem = Invoker.PropertyGet(this, "RowCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowCount
		/// </summary>
		/// <param name="section">Int16 Section</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 RowCount(Int16 section)
		{
			return get_RowCount(section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowsCellCount(Int16 section, Int16 row)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, row);
			object returnItem = Invoker.PropertyGet(this, "RowsCellCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowsCellCount
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 RowsCellCount(Int16 section, Int16 row)
		{
			return get_RowsCellCount(section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowType(Int16 section, Int16 row)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, row);
			object returnItem = Invoker.PropertyGet(this, "RowType", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_RowType(Int16 section, Int16 row, Int16 value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, row);
			Invoker.PropertySet(this, "RowType", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowType
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 RowType(Int16 section, Int16 row)
		{
			return get_RowType(section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVConnects Connects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connects", paramsArray);
				NetOffice.VisioApi.IVConnects newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVConnects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 Index16
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index16", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Style", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string StyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "StyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "StyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string LineStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LineStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LineStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string LineStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LineStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LineStyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FillStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FillStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FillStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FillStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FillStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FillStyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string TextStyle
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyle", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextStyle", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string TextStyleKeepFmt
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextStyleKeepFmt", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TextStyleKeepFmt", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double old_AreaIU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_AreaIU", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double old_LengthIU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "old_LengthIU", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="fFill">Int16 fFill</param>
		/// <param name="lineRes">Double LineRes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_GeomExIf(Int16 fFill, Double lineRes)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fFill, lineRes);
			object returnItem = Invoker.PropertyGet(this, "GeomExIf", paramsArray);
			ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_GeomExIf
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="fFill">Int16 fFill</param>
		/// <param name="lineRes">Double LineRes</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object GeomExIf(Int16 fFill, Double lineRes)
		{
			return get_GeomExIf(fFill, lineRes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_UniqueID(Int16 fUniqueID)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(fUniqueID);
			object returnItem = Invoker.PropertyGet(this, "UniqueID", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_UniqueID
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string UniqueID(Int16 fUniqueID)
		{
			return get_UniqueID(fUniqueID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingPage", paramsArray);
				NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingMaster", paramsArray);
				NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingShape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_SectionExists(Int16 section, Int16 fExistsLocally)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, fExistsLocally);
			object returnItem = Invoker.PropertyGet(this, "SectionExists", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SectionExists
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SectionExists(Int16 section, Int16 fExistsLocally)
		{
			return get_SectionExists(section, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, row, fExistsLocally);
			object returnItem = Invoker.PropertyGet(this, "RowExists", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowExists
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
		{
			return get_RowExists(section, row, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeSpecificCellName, fExistsLocally);
			object returnItem = Invoker.PropertyGet(this, "CellExists", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExists
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally)
		{
			return get_CellExists(localeSpecificCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="column">Int16 Column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(section, row, column, fExistsLocally);
			object returnItem = Invoker.PropertyGet(this, "CellsSRCExists", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRCExists
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="column">Int16 Column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{
			return get_CellsSRCExists(section, row, column, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 LayerCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LayerCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVLayer get_Layer(Int16 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Layer", paramsArray);
			NetOffice.VisioApi.IVLayer newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVLayer;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Layer
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVLayer Layer(Int16 index)
		{
			return get_Layer(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EventList", paramsArray);
				NetOffice.VisioApi.IVEventList newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVEventList;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PersistsEvents", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string ClassID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ClassID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ForeignType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ForeignType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object Object
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Object", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ID16
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID16", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVConnects FromConnects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FromConnects", paramsArray);
				NetOffice.VisioApi.IVConnects newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVConnects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVHyperlink Hyperlink
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlink", paramsArray);
				NetOffice.VisioApi.IVHyperlink newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVHyperlink;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string ProgID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProgID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ObjectIsInherited
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ObjectIsInherited", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPaths Paths
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Paths", paramsArray);
				NetOffice.VisioApi.IVPaths newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPaths;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVPaths PathsLocal
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PathsLocal", paramsArray);
				NetOffice.VisioApi.IVPaths newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPaths;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSection get_Section(Int16 index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Section", paramsArray);
			NetOffice.VisioApi.IVSection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <param name="index">Int16 Index</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSection Section(Int16 index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVHyperlinks Hyperlinks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Hyperlinks", paramsArray);
				NetOffice.VisioApi.IVHyperlinks newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVHyperlinks;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(otherShape, tolerance, flags);
			object returnItem = Invoker.PropertyGet(this, "SpatialRelation", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialRelation
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
		{
			return get_SpatialRelation(otherShape, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(otherShape, flags);
			object returnItem = Invoker.PropertyGet(this, "DistanceFrom", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFrom
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags)
		{
			return get_DistanceFrom(otherShape, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		/// <param name="pvt">optional object pvt</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, flags, pvPathIndex, pvCurveIndex, pvt);
			object returnItem = Invoker.PropertyGet(this, "DistanceFromPoint", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		/// <param name="pvt">optional object pvt</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt)
		{
			return get_DistanceFromPoint(x, y, flags, pvPathIndex, pvCurveIndex, pvt);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, flags);
			object returnItem = Invoker.PropertyGet(this, "DistanceFromPoint", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double DistanceFromPoint(Double x, Double y, Int16 flags)
		{
			return get_DistanceFromPoint(x, y, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, flags, pvPathIndex);
			object returnItem = Invoker.PropertyGet(this, "DistanceFromPoint", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex)
		{
			return get_DistanceFromPoint(x, y, flags, pvPathIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, flags, pvPathIndex, pvCurveIndex);
			object returnItem = Invoker.PropertyGet(this, "DistanceFromPoint", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
		{
			return get_DistanceFromPoint(x, y, flags, pvPathIndex, pvCurveIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="resultRoot">optional object ResultRoot</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(relation, tolerance, flags, resultRoot);
			object returnItem = Invoker.PropertyGet(this, "SpatialNeighbors", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialNeighbors
		/// </summary>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="resultRoot">optional object ResultRoot</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
		{
			return get_SpatialNeighbors(relation, tolerance, flags, resultRoot);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(relation, tolerance, flags);
			object returnItem = Invoker.PropertyGet(this, "SpatialNeighbors", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialNeighbors
		/// </summary>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags)
		{
			return get_SpatialNeighbors(relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, relation, tolerance, flags);
			object returnItem = Invoker.PropertyGet(this, "SpatialSearch", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialSearch
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 Relation</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return get_SpatialSearch(x, y, relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_CellsU(string localeIndependentCellName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeIndependentCellName);
			object returnItem = Invoker.PropertyGet(this, "CellsU", paramsArray);
			NetOffice.VisioApi.IVCell newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell CellsU(string localeIndependentCellName)
		{
			return get_CellsU(localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string NameU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NameU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "NameU", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeIndependentCellName, fExistsLocally);
			object returnItem = Invoker.PropertyGet(this, "CellExistsU", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExistsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{
			return get_CellExistsU(localeIndependentCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsRowIndex(string localeSpecificCellName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeSpecificCellName);
			object returnItem = Invoker.PropertyGet(this, "CellsRowIndex", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsRowIndex
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CellsRowIndex(string localeSpecificCellName)
		{
			return get_CellsRowIndex(localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsRowIndexU(string localeIndependentCellName)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(localeIndependentCellName);
			object returnItem = Invoker.PropertyGet(this, "CellsRowIndexU", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsRowIndexU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CellsRowIndexU(string localeIndependentCellName)
		{
			return get_CellsRowIndexU(localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public bool IsOpenForTextEdit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsOpenForTextEdit", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape RootShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RootShape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape MasterShape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MasterShape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
				stdole.Picture newObject = Factory.CreateObjectFromComProxy(this,returnItem) as stdole.Picture;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public byte[] ForeignData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "ForeignData", paramsArray);
				return (byte[])returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 Language
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Language", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Language", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double AreaIU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AreaIU", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double LengthIU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LengthIU", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ContainingPageID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingPageID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ContainingMasterID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingMasterID", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public NetOffice.VisioApi.IVMaster DataGraphic
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataGraphic", paramsArray);
				NetOffice.VisioApi.IVMaster newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVMaster;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataGraphic", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public bool IsDataGraphicCallout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDataGraphicCallout", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVContainerProperties ContainerProperties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainerProperties", paramsArray);
				NetOffice.VisioApi.IVContainerProperties newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVContainerProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] MemberOfContainers
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "MemberOfContainers", paramsArray);
				return (Int32[])returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool IsCallout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsCallout", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVShape CalloutTarget
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CalloutTarget", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CalloutTarget", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] CalloutsAssociated
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = (object)Invoker.PropertyGet(this, "CalloutsAssociated", paramsArray);
				return (Int32[])returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVComments Comments
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Comments", paramsArray);
				NetOffice.VisioApi.IVComments newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVComments;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void VoidGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "VoidGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BringForward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringForward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BringToFront()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringToFront", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ConvertToGroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ConvertToGroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FlipHorizontal()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FlipHorizontal", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FlipVertical()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "FlipVertical", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void ReverseEnds()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ReverseEnds", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SendBackward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendBackward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SendToBack()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendToBack", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Rotate90()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Rotate90", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Ungroup()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Ungroup", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "old_Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void old_Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "old_Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void VoidDuplicate()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "VoidDuplicate", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(objectToDrop, xPos, yPos);
			object returnItem = Invoker.MethodReturn(this, "Drop", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 AddSection(Int16 section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section);
			object returnItem = Invoker.MethodReturn(this, "AddSection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void DeleteSection(Int16 section)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section);
			Invoker.Method(this, "DeleteSection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="rowTag">Int16 RowTag</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 AddRow(Int16 section, Int16 row, Int16 rowTag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, row, rowTag);
			object returnItem = Invoker.MethodReturn(this, "AddRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void DeleteRow(Int16 section, Int16 row)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, row);
			Invoker.Method(this, "DeleteRow", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetCenter(Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPos, yPos);
			Invoker.Method(this, "SetCenter", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetBegin(Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPos, yPos);
			Invoker.Method(this, "SetBegin", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SetEnd(Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPos, yPos);
			Invoker.Method(this, "SetEnd", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Export(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "Export", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="rowName">string RowName</param>
		/// <param name="rowTag">Int16 RowTag</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, rowName, rowTag);
			object returnItem = Invoker.MethodReturn(this, "AddNamedRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="section">Int16 Section</param>
		/// <param name="row">Int16 Row</param>
		/// <param name="rowTag">Int16 RowTag</param>
		/// <param name="rowCount">Int16 RowCount</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(section, row, rowTag, rowCount);
			object returnItem = Invoker.MethodReturn(this, "AddRows", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd);
			object returnItem = Invoker.MethodReturn(this, "DrawLine", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2);
			object returnItem = Invoker.MethodReturn(this, "DrawRectangle", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(x1, y1, x2, y2);
			object returnItem = Invoker.MethodReturn(this, "DrawOval", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, tolerance, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawSpline", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, degree, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawBezier", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, flags);
			object returnItem = Invoker.MethodReturn(this, "DrawPolyline", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="tolerance">Double Tolerance</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void FitCurve(Double tolerance, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(tolerance, flags);
			Invoker.Method(this, "FitCurve", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Import(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "Import", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void CenterDrawing()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CenterDrawing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, flags);
			object returnItem = Invoker.MethodReturn(this, "InsertFromFile", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="classOrProgID">string ClassOrProgID</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(classOrProgID, flags);
			object returnItem = Invoker.MethodReturn(this, "InsertObject", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow OpenDrawWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OpenDrawWindow", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVWindow OpenSheetWindow()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OpenSheetWindow", paramsArray);
			NetOffice.VisioApi.IVWindow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVWindow;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectsToInstance">object[] ObjectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] IDArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 DropMany(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			iDArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
			object returnItem = Invoker.MethodReturn(this, "DropMany", paramsArray);
			iDArray = (Int16[])paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sRCStream">Int16[] SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetFormulas(Int16[] sRCStream, out object[] formulaArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			formulaArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray);
			Invoker.Method(this, "GetFormulas", paramsArray, modifiers);
			formulaArray = (object[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sRCStream">Int16[] SRCStream</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="unitsNamesOrCodes">object[] UnitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetResults(Int16[] sRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			resultArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, flags, (object)unitsNamesOrCodes, (object)resultArray);
			Invoker.Method(this, "GetResults", paramsArray, modifiers);
			resultArray = (object[])paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sRCStream">Int16[] SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray, flags);
			object returnItem = Invoker.MethodReturn(this, "SetFormulas", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sRCStream">Int16[] SRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] UnitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
			object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Layout()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Layout", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true,true);
			lpr8Left = 0;
			lpr8Bottom = 0;
			lpr8Right = 0;
			lpr8Top = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(flags, lpr8Left, lpr8Bottom, lpr8Right, lpr8Top);
			Invoker.Method(this, "BoundingBox", paramsArray, modifiers);
			lpr8Left = (Double)paramsArray[1];
			lpr8Bottom = (Double)paramsArray[2];
			lpr8Right = (Double)paramsArray[3];
			lpr8Top = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="tolerance">Double Tolerance</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 HitTest(Double xPos, Double yPos, Double tolerance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xPos, yPos, tolerance);
			object returnItem = Invoker.MethodReturn(this, "HitTest", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVHyperlink AddHyperlink()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "AddHyperlink", paramsArray);
			NetOffice.VisioApi.IVHyperlink newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVHyperlink;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void TransformXYTo(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			xprime = 0;
			yprime = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(otherShape, x, y, xprime, yprime);
			Invoker.Method(this, "TransformXYTo", paramsArray, modifiers);
			xprime = (Double)paramsArray[3];
			yprime = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape OtherShape</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void TransformXYFrom(NetOffice.VisioApi.IVShape otherShape, Double x, Double y, out Double xprime, out Double yprime)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true,true);
			xprime = 0;
			yprime = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(otherShape, x, y, xprime, yprime);
			Invoker.Method(this, "TransformXYFrom", paramsArray, modifiers);
			xprime = (Double)paramsArray[3];
			yprime = (Double)paramsArray[4];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void XYToPage(Double x, Double y, out Double xprime, out Double yprime)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			xprime = 0;
			yprime = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, xprime, yprime);
			Invoker.Method(this, "XYToPage", paramsArray, modifiers);
			xprime = (Double)paramsArray[2];
			yprime = (Double)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void XYFromPage(Double x, Double y, out Double xprime, out Double yprime)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true,true);
			xprime = 0;
			yprime = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(x, y, xprime, yprime);
			Invoker.Method(this, "XYFromPage", paramsArray, modifiers);
			xprime = (Double)paramsArray[2];
			yprime = (Double)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void UpdateAlignmentBox()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateAlignmentBox", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="objectsToInstance">object[] ObjectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] IDArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 DropManyU(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			iDArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)objectsToInstance, (object)xyArray, (object)iDArray);
			object returnItem = Invoker.MethodReturn(this, "DropManyU", paramsArray);
			iDArray = (Int16[])paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sRCStream">Int16[] SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GetFormulasU(Int16[] sRCStream, out object[] formulaArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			formulaArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray);
			Invoker.Method(this, "GetFormulasU", paramsArray, modifiers);
			formulaArray = (object[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		/// <param name="weights">optional object weights</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots, weights);
			object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 Flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots);
			object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Group()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Group", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Duplicate()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Duplicate", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void SwapEnds()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SwapEnds", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Copy(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Cut(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">optional object Flags</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Paste(object flags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags);
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Paste", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		/// <param name="link">optional object Link</param>
		/// <param name="displayAsIcon">optional object DisplayAsIcon</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link, object displayAsIcon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, link, displayAsIcon);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="format">Int32 Format</param>
		/// <param name="link">optional object Link</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, link);
			Invoker.Method(this, "PasteSpecial", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		/// <param name="data">optional object Data</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType, iterationMode, data);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes SelType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(selType, iterationMode);
			object returnItem = Invoker.MethodReturn(this, "CreateSelection", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="distance">Double Distance</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Offset(Double distance)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(distance);
			Invoker.Method(this, "Offset", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="type">Int16 Type</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(type, xPos, yPos);
			object returnItem = Invoker.MethodReturn(this, "AddGuide", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="xControl">Double xControl</param>
		/// <param name="yControl">Double yControl</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd, xControl, yControl);
			object returnItem = Invoker.MethodReturn(this, "DrawArcByThreePoints", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags SweepFlag</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xBegin, yBegin, xEnd, yEnd, sweepFlag);
			object returnItem = Invoker.MethodReturn(this, "DrawQuarterArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		/// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius, startAngle, endAngle);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double Radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(xCenter, yCenter, radius, startAngle);
			object returnItem = Invoker.MethodReturn(this, "DrawCircularArc", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="rowID">Int32 RowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 rowID, object applyDataGraphicAfterLink)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, rowID, applyDataGraphicAfterLink);
			Invoker.Method(this, "LinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="rowID">Int32 RowID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 rowID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, rowID);
			Invoker.Method(this, "LinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void BreakLinkToData(Int32 dataRecordsetID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID);
			Invoker.Method(this, "BreakLinkToData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public Int32 GetLinkedDataRow(Int32 dataRecordsetID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID);
			object returnItem = Invoker.MethodReturn(this, "GetLinkedDataRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetIDs">Int32[] DataRecordsetIDs</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetLinkedDataRecordsetIDs(out Int32[] dataRecordsetIDs)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			dataRecordsetIDs = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)dataRecordsetIDs);
			Invoker.Method(this, "GetLinkedDataRecordsetIDs", paramsArray, modifiers);
			dataRecordsetIDs = (Int32[])paramsArray[0];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="customPropertyIndices">Int32[] CustomPropertyIndices</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void GetCustomPropertiesLinkedToData(Int32 dataRecordsetID, out Int32[] customPropertyIndices)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			customPropertyIndices = null;
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, (object)customPropertyIndices);
			Invoker.Method(this, "GetCustomPropertiesLinkedToData", paramsArray, modifiers);
			customPropertyIndices = (Int32[])paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="customPropertyIndex">Int32 CustomPropertyIndex</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public bool IsCustomPropertyLinked(Int32 dataRecordsetID, Int32 customPropertyIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, customPropertyIndex);
			object returnItem = Invoker.MethodReturn(this, "IsCustomPropertyLinked", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="dataRecordsetID">Int32 DataRecordsetID</param>
		/// <param name="customPropertyIndex">Int32 CustomPropertyIndex</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public string GetCustomPropertyLinkedColumn(Int32 dataRecordsetID, Int32 customPropertyIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataRecordsetID, customPropertyIndex);
			object returnItem = Invoker.MethodReturn(this, "GetCustomPropertyLinkedColumn", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="toShape">NetOffice.VisioApi.IVShape ToShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir PlacementDir</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(toShape, placementDir, connector);
			Invoker.Method(this, "AutoConnect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="toShape">NetOffice.VisioApi.IVShape ToShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir PlacementDir</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(toShape, placementDir);
			Invoker.Method(this, "AutoConnect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="category">string Category</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public bool HasCategory(string category)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category);
			object returnItem = Invoker.MethodReturn(this, "HasCategory", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisConnectedShapesFlags Flags</param>
		/// <param name="categoryFilter">string CategoryFilter</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] ConnectedShapes(NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags, string categoryFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
			object returnItem = (object)Invoker.MethodReturn(this, "ConnectedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags Flags</param>
		/// <param name="categoryFilter">string CategoryFilter</param>
		/// <param name="pOtherConnectedShape">optional NetOffice.VisioApi.IVShape pOtherConnectedShape</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter, object pOtherConnectedShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter, pOtherConnectedShape);
			object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags Flags</param>
		/// <param name="categoryFilter">string CategoryFilter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
			object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="connectorEnd">NetOffice.VisioApi.Enums.VisConnectorEnds ConnectorEnd</param>
		/// <param name="offsetX">Double OffsetX</param>
		/// <param name="offsetY">Double OffsetY</param>
		/// <param name="units">NetOffice.VisioApi.Enums.VisUnitCodes Units</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void Disconnect(NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd, Double offsetX, Double offsetY, NetOffice.VisioApi.Enums.VisUnitCodes units)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectorEnd, offsetX, offsetY, units);
			Invoker.Method(this, "Disconnect", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection Direction</param>
		/// <param name="distance">Double Distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes UnitCode</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(direction, distance, unitCode);
			Invoker.Method(this, "Resize", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void AddToContainers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AddToContainers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void RemoveFromContainers()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RemoveFromContainers", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVPage CreateSubProcess()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateSubProcess", paramsArray);
			NetOffice.VisioApi.IVPage newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVPage;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage Page</param>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(page, objectToDrop, newShape);
			object returnItem = Invoker.MethodReturn(this, "MoveToSubprocess", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage Page</param>
		/// <param name="objectToDrop">object ObjectToDrop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(page, objectToDrop);
			object returnItem = Invoker.MethodReturn(this, "MoveToSubprocess", paramsArray);
			NetOffice.VisioApi.IVSelection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVSelection;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// 
		/// </summary>
		/// <param name="delFlags">Int32 DelFlags</param>
		[SupportByVersionAttribute("Visio", 14,15,16)]
		public void DeleteEx(Int32 delFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(delFlags);
			Invoker.Method(this, "DeleteEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object MasterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop, replaceFlags);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object MasterOrMasterShortcutToDrop</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(masterOrMasterShortcutToDrop);
			object returnItem = Invoker.MethodReturn(this, "ReplaceShape", paramsArray);
			NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor);
			Invoker.Method(this, "SetQuickStyle", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="changePictureFlags">optional Int32 ChangePictureFlags = 0</param>
		[SupportByVersionAttribute("Visio", 15, 16)]
		public Double ChangePicture(string fileName, object changePictureFlags)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, changePictureFlags);
			object returnItem = Invoker.MethodReturn(this, "ChangePicture", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 15, 16)]
		public Double ChangePicture(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			object returnItem = Invoker.MethodReturn(this, "ChangePicture", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}