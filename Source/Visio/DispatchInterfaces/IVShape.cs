using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// IVShape
	/// </summary>
	[SyntaxBypass]
 	public class IVShape_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVShape_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVShape_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_AreaIU(object fIncludeSubShapes)
		{
			return Factory.ExecuteDoublePropertyGet(this, "AreaIU", fIncludeSubShapes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_AreaIU
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_AreaIU")]
		public Double AreaIU(object fIncludeSubShapes)
		{
			return get_AreaIU(fIncludeSubShapes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_LengthIU(object fIncludeSubShapes)
		{
			return Factory.ExecuteDoublePropertyGet(this, "LengthIU", fIncludeSubShapes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_LengthIU
		/// </summary>
		/// <param name="fIncludeSubShapes">optional bool fIncludeSubShapes</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_LengthIU")]
		public Double LengthIU(object fIncludeSubShapes)
		{
			return get_LengthIU(fIncludeSubShapes);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface IVShape 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVShape : IVShape_
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
                    _type = typeof(IVShape);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVShape(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVShape(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVShape(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Parent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Stat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster Master
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "Master");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Type
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "Cells", NetOffice.VisioApi.IVCell.LateBindingApiWrapperType, localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Cells")]
		public NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName)
		{
			return get_Cells(localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsSRC", NetOffice.VisioApi.IVCell.LateBindingApiWrapperType, section, row, column);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRC
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsSRC")]
		public NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column)
		{
			return get_CellsSRC(section, row, column);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShapes Shapes
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShapes>(this, "Shapes");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Data1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Data1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Data1", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Data2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Data2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Data2", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Data3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Data3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Data3", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Help
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Help");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Help", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string NameID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 CharCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CharCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVCharacters Characters
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCharacters>(this, "Characters");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 OneD
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "OneD");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OneD", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 GeometryCount
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "GeometryCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowCount(Int16 section)
		{
			return Factory.ExecuteInt16PropertyGet(this, "RowCount", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowCount")]
		public Int16 RowCount(Int16 section)
		{
			return get_RowCount(section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowsCellCount(Int16 section, Int16 row)
		{
			return Factory.ExecuteInt16PropertyGet(this, "RowsCellCount", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowsCellCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowsCellCount")]
		public Int16 RowsCellCount(Int16 section, Int16 row)
		{
			return get_RowsCellCount(section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowType(Int16 section, Int16 row)
		{
			return Factory.ExecuteInt16PropertyGet(this, "RowType", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_RowType(Int16 section, Int16 row, Int16 value)
		{
			Factory.ExecutePropertySet(this, "RowType", section, row, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowType
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowType")]
		public Int16 RowType(Int16 section, Int16 row)
		{
			return get_RowType(section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVConnects Connects
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "Connects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 Index16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Index16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Style
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Style");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string StyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "StyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string LineStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LineStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string LineStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LineStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FillStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FillStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FillStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FillStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FillStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TextStyle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string TextStyleKeepFmt
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextStyleKeepFmt");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double old_AreaIU
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "old_AreaIU");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double old_LengthIU
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "old_LengthIU");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="fFill">Int16 fFill</param>
		/// <param name="lineRes">Double lineRes</param>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_GeomExIf(Int16 fFill, Double lineRes)
		{
			return Factory.ExecuteReferencePropertyGet(this, "GeomExIf", fFill, lineRes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_GeomExIf
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="fFill">Int16 fFill</param>
		/// <param name="lineRes">Double lineRes</param>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult, Redirect("get_GeomExIf")]
		public object GeomExIf(Int16 fFill, Double lineRes)
		{
			return get_GeomExIf(fFill, lineRes);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_UniqueID(Int16 fUniqueID)
		{
			return Factory.ExecuteStringPropertyGet(this, "UniqueID", fUniqueID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_UniqueID
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_UniqueID")]
		public string UniqueID(Int16 fUniqueID)
		{
			return get_UniqueID(fUniqueID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ContainingPage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "ContainingMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "ContainingShape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_SectionExists(Int16 section, Int16 fExistsLocally)
		{
			return Factory.ExecuteInt16PropertyGet(this, "SectionExists", section, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SectionExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SectionExists")]
		public Int16 SectionExists(Int16 section, Int16 fExistsLocally)
		{
			return get_SectionExists(section, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
		{
			return Factory.ExecuteInt16PropertyGet(this, "RowExists", section, row, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowExists")]
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CellExists", localeSpecificCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExists
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellExists")]
		public Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally)
		{
			return get_CellExists(localeSpecificCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CellsSRCExists", section, row, column, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRCExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsSRCExists")]
		public Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{
			return get_CellsSRCExists(section, row, column, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 LayerCount
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "LayerCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVLayer get_Layer(Int16 index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVLayer>(this, "Layer", NetOffice.VisioApi.IVLayer.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Layer
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Layer")]
		public NetOffice.VisioApi.IVLayer Layer(Int16 index)
		{
			return get_Layer(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 PersistsEvents
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string ClassID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ClassID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ForeignType
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ForeignType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public object Object
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 ID16
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ID16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVConnects FromConnects
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "FromConnects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVHyperlink Hyperlink
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVHyperlink>(this, "Hyperlink");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string ProgID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ProgID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 ObjectIsInherited
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "ObjectIsInherited");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPaths Paths
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPaths>(this, "Paths");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPaths PathsLocal
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPaths>(this, "PathsLocal");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSection get_Section(Int16 index)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSection>(this, "Section", NetOffice.VisioApi.IVSection.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Section")]
		public NetOffice.VisioApi.IVSection Section(Int16 index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVHyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVHyperlinks>(this, "Hyperlinks");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
		{
			return Factory.ExecuteInt16PropertyGet(this, "SpatialRelation", otherShape, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialRelation
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialRelation")]
		public Int16 SpatialRelation(NetOffice.VisioApi.IVShape otherShape, Double tolerance, Int16 flags)
		{
			return get_SpatialRelation(otherShape, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFrom(NetOffice.VisioApi.IVShape otherShape, Int16 flags)
		{
			return Factory.ExecuteDoublePropertyGet(this, "DistanceFrom", otherShape, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFrom
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_DistanceFrom")]
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
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		/// <param name="pvt">optional object pvt</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex, object pvt)
		{
			return Factory.ExecuteDoublePropertyGet(this, "DistanceFromPoint", new object[]{ x, y, flags, pvPathIndex, pvCurveIndex, pvt });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		/// <param name="pvt">optional object pvt</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_DistanceFromPoint")]
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
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags)
		{
			return Factory.ExecuteDoublePropertyGet(this, "DistanceFromPoint", x, y, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_DistanceFromPoint")]
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
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex)
		{
			return Factory.ExecuteDoublePropertyGet(this, "DistanceFromPoint", x, y, flags, pvPathIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_DistanceFromPoint")]
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
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
		{
			return Factory.ExecuteDoublePropertyGet(this, "DistanceFromPoint", new object[]{ x, y, flags, pvPathIndex, pvCurveIndex });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_DistanceFromPoint
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="pvPathIndex">optional object pvPathIndex</param>
		/// <param name="pvCurveIndex">optional object pvCurveIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_DistanceFromPoint")]
		public Double DistanceFromPoint(Double x, Double y, Int16 flags, object pvPathIndex, object pvCurveIndex)
		{
			return get_DistanceFromPoint(x, y, flags, pvPathIndex, pvCurveIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="resultRoot">optional object resultRoot</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialNeighbors", NetOffice.VisioApi.IVSelection.LateBindingApiWrapperType, relation, tolerance, flags, resultRoot);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialNeighbors
		/// </summary>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="resultRoot">optional object resultRoot</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialNeighbors")]
		public NetOffice.VisioApi.IVSelection SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags, object resultRoot)
		{
			return get_SpatialNeighbors(relation, tolerance, flags, resultRoot);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialNeighbors(Int16 relation, Double tolerance, Int16 flags)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialNeighbors", NetOffice.VisioApi.IVSelection.LateBindingApiWrapperType, relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialNeighbors
		/// </summary>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialNeighbors")]
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
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSelection>(this, "SpatialSearch", NetOffice.VisioApi.IVSelection.LateBindingApiWrapperType, new object[]{ x, y, relation, tolerance, flags });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialSearch
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialSearch")]
		public NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags)
		{
			return get_SpatialSearch(x, y, relation, tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.VisioApi.IVCell get_CellsU(string localeIndependentCellName)
		{
			return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsU", NetOffice.VisioApi.IVCell.LateBindingApiWrapperType, localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsU")]
		public NetOffice.VisioApi.IVCell CellsU(string localeIndependentCellName)
		{
			return get_CellsU(localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string NameU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CellExistsU", localeIndependentCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExistsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellExistsU")]
		public Int16 CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{
			return get_CellExistsU(localeIndependentCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsRowIndex(string localeSpecificCellName)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CellsRowIndex", localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsRowIndex
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsRowIndex")]
		public Int16 CellsRowIndex(string localeSpecificCellName)
		{
			return get_CellsRowIndex(localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CellsRowIndexU(string localeIndependentCellName)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CellsRowIndexU", localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsRowIndexU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsRowIndexU")]
		public Int16 CellsRowIndexU(string localeIndependentCellName)
		{
			return get_CellsRowIndexU(localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public bool IsOpenForTextEdit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsOpenForTextEdit");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape RootShape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "RootShape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape MasterShape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "MasterShape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		public stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
            }
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Language
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Language");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Language", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double AreaIU
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "AreaIU");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double LengthIU
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "LengthIU");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ContainingPageID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContainingPageID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 ContainingMasterID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContainingMasterID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVMaster DataGraphic
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "DataGraphic");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataGraphic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool IsDataGraphicCallout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDataGraphicCallout");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVContainerProperties ContainerProperties
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVContainerProperties>(this, "ContainerProperties");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
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
		[SupportByVersion("Visio", 14,15,16)]
		public bool IsCallout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsCallout");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape CalloutTarget
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "CalloutTarget");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "CalloutTarget", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
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
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		public NetOffice.VisioApi.IVComments Comments
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVComments>(this, "Comments");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void VoidGroup()
		{
			 Factory.ExecuteMethod(this, "VoidGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void BringForward()
		{
			 Factory.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void BringToFront()
		{
			 Factory.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ConvertToGroup()
		{
			 Factory.ExecuteMethod(this, "ConvertToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FlipHorizontal()
		{
			 Factory.ExecuteMethod(this, "FlipHorizontal");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FlipVertical()
		{
			 Factory.ExecuteMethod(this, "FlipVertical");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void ReverseEnds()
		{
			 Factory.ExecuteMethod(this, "ReverseEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SendBackward()
		{
			 Factory.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SendToBack()
		{
			 Factory.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Rotate90()
		{
			 Factory.ExecuteMethod(this, "Rotate90");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Ungroup()
		{
			 Factory.ExecuteMethod(this, "Ungroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_Copy()
		{
			 Factory.ExecuteMethod(this, "old_Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void old_Cut()
		{
			 Factory.ExecuteMethod(this, "old_Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void VoidDuplicate()
		{
			 Factory.ExecuteMethod(this, "VoidDuplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Drop", objectToDrop, xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AddSection(Int16 section)
		{
			return Factory.ExecuteInt16MethodGet(this, "AddSection", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DeleteSection(Int16 section)
		{
			 Factory.ExecuteMethod(this, "DeleteSection", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AddRow(Int16 section, Int16 row, Int16 rowTag)
		{
			return Factory.ExecuteInt16MethodGet(this, "AddRow", section, row, rowTag);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void DeleteRow(Int16 section, Int16 row)
		{
			 Factory.ExecuteMethod(this, "DeleteRow", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetCenter(Double xPos, Double yPos)
		{
			 Factory.ExecuteMethod(this, "SetCenter", xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetBegin(Double xPos, Double yPos)
		{
			 Factory.ExecuteMethod(this, "SetBegin", xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SetEnd(Double xPos, Double yPos)
		{
			 Factory.ExecuteMethod(this, "SetEnd", xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Export(string fileName)
		{
			 Factory.ExecuteMethod(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="rowName">string rowName</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag)
		{
			return Factory.ExecuteInt16MethodGet(this, "AddNamedRow", section, rowName, rowTag);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		/// <param name="rowCount">Int16 rowCount</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount)
		{
			return Factory.ExecuteInt16MethodGet(this, "AddRows", section, row, rowTag, rowCount);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawLine", xBegin, yBegin, xEnd, yEnd);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawRectangle", x1, y1, x2, y2);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawOval", x1, y1, x2, y2);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, tolerance, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawSpline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, degree, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawBezier", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)xyArray, flags);
            object returnItem = Invoker.MethodReturn(this, "DrawPolyline", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void FitCurve(Double tolerance, Int16 flags)
		{
			 Factory.ExecuteMethod(this, "FitCurve", tolerance, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Import(string fileName)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Import", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void CenterDrawing()
		{
			 Factory.ExecuteMethod(this, "CenterDrawing");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertFromFile", fileName, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classOrProgID">string classOrProgID</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "InsertObject", classOrProgID, flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow OpenDrawWindow()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenDrawWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVWindow OpenSheetWindow()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenSheetWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetFormulas", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Layout()
		{
			 Factory.ExecuteMethod(this, "Layout");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="tolerance">Double tolerance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 HitTest(Double xPos, Double yPos, Double tolerance)
		{
			return Factory.ExecuteInt16MethodGet(this, "HitTest", xPos, yPos, tolerance);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVHyperlink AddHyperlink()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVHyperlink>(this, "AddHyperlink");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="otherShape">NetOffice.VisioApi.IVShape otherShape</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="xprime">Double xprime</param>
		/// <param name="yprime">Double yprime</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void UpdateAlignmentBox()
		{
			 Factory.ExecuteMethod(this, "UpdateAlignmentBox");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		/// <param name="weights">optional object weights</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots, weights);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots)
		{
            object[] paramsArray = Invoker.ValidateParamsArray(degree, flags, (object)xyArray, (object)knots);
            object returnItem = Invoker.MethodReturn(this, "DrawNURBS", paramsArray);
            NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this, returnItem, false) as NetOffice.VisioApi.IVShape;
            return newObject;
        }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Group()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Duplicate()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Duplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void SwapEnds()
		{
			 Factory.ExecuteMethod(this, "SwapEnds");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy(object flags)
		{
			 Factory.ExecuteMethod(this, "Copy", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut(object flags)
		{
			 Factory.ExecuteMethod(this, "Cut", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste(object flags)
		{
			 Factory.ExecuteMethod(this, "Paste", flags);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void PasteSpecial(Int32 format, object link)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", format, link);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		/// <param name="data">optional object data</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode, data);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "CreateSelection", selType, iterationMode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="distance">Double distance</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Offset(Double distance)
		{
			 Factory.ExecuteMethod(this, "Offset", distance);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "AddGuide", type, xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="xControl">Double xControl</param>
		/// <param name="yControl">Double yControl</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawArcByThreePoints", new object[]{ xBegin, yBegin, xEnd, yEnd, xControl, yControl });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawQuarterArc", new object[]{ xBegin, yBegin, xEnd, yEnd, sweepFlag });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		/// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", new object[]{ xCenter, yCenter, radius, startAngle, endAngle });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "DrawCircularArc", xCenter, yCenter, radius, startAngle);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="rowID">Int32 rowID</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 rowID, object applyDataGraphicAfterLink)
		{
			 Factory.ExecuteMethod(this, "LinkToData", dataRecordsetID, rowID, applyDataGraphicAfterLink);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="rowID">Int32 rowID</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		public void LinkToData(Int32 dataRecordsetID, Int32 rowID)
		{
			 Factory.ExecuteMethod(this, "LinkToData", dataRecordsetID, rowID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void BreakLinkToData(Int32 dataRecordsetID)
		{
			 Factory.ExecuteMethod(this, "BreakLinkToData", dataRecordsetID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public Int32 GetLinkedDataRow(Int32 dataRecordsetID)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetLinkedDataRow", dataRecordsetID);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetIDs">Int32[] dataRecordsetIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="customPropertyIndices">Int32[] customPropertyIndices</param>
		[SupportByVersion("Visio", 12,14,15,16)]
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
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public bool IsCustomPropertyLinked(Int32 dataRecordsetID, Int32 customPropertyIndex)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsCustomPropertyLinked", dataRecordsetID, customPropertyIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="customPropertyIndex">Int32 customPropertyIndex</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public string GetCustomPropertyLinkedColumn(Int32 dataRecordsetID, Int32 customPropertyIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCustomPropertyLinkedColumn", dataRecordsetID, customPropertyIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector)
		{
			 Factory.ExecuteMethod(this, "AutoConnect", toShape, placementDir, connector);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="toShape">NetOffice.VisioApi.IVShape toShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		public void AutoConnect(NetOffice.VisioApi.IVShape toShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir)
		{
			 Factory.ExecuteMethod(this, "AutoConnect", toShape, placementDir);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="category">string category</param>
		[SupportByVersion("Visio", 14,15,16)]
		public bool HasCategory(string category)
		{
			return Factory.ExecuteBoolMethodGet(this, "HasCategory", category);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags</param>
		/// <param name="categoryFilter">string categoryFilter</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] ConnectedShapes(NetOffice.VisioApi.Enums.VisConnectedShapesFlags flags, string categoryFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
			object returnItem = (object)Invoker.MethodReturn(this, "ConnectedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
		/// <param name="categoryFilter">string categoryFilter</param>
		/// <param name="pOtherConnectedShape">optional NetOffice.VisioApi.IVShape pOtherConnectedShape</param>
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter, object pOtherConnectedShape)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter, pOtherConnectedShape);
			object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="flags">NetOffice.VisioApi.Enums.VisGluedShapesFlags flags</param>
		/// <param name="categoryFilter">string categoryFilter</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		public Int32[] GluedShapes(NetOffice.VisioApi.Enums.VisGluedShapesFlags flags, string categoryFilter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(flags, categoryFilter);
			object returnItem = (object)Invoker.MethodReturn(this, "GluedShapes", paramsArray);
			return (Int32[])returnItem;
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="connectorEnd">NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd</param>
		/// <param name="offsetX">Double offsetX</param>
		/// <param name="offsetY">Double offsetY</param>
		/// <param name="units">NetOffice.VisioApi.Enums.VisUnitCodes units</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void Disconnect(NetOffice.VisioApi.Enums.VisConnectorEnds connectorEnd, Double offsetX, Double offsetY, NetOffice.VisioApi.Enums.VisUnitCodes units)
		{
			 Factory.ExecuteMethod(this, "Disconnect", connectorEnd, offsetX, offsetY, units);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisResizeDirection direction</param>
		/// <param name="distance">Double distance</param>
		/// <param name="unitCode">NetOffice.VisioApi.Enums.VisUnitCodes unitCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void Resize(NetOffice.VisioApi.Enums.VisResizeDirection direction, Double distance, NetOffice.VisioApi.Enums.VisUnitCodes unitCode)
		{
			 Factory.ExecuteMethod(this, "Resize", direction, distance, unitCode);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void AddToContainers()
		{
			 Factory.ExecuteMethod(this, "AddToContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		public void RemoveFromContainers()
		{
			 Factory.ExecuteMethod(this, "RemoveFromContainers");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVPage CreateSubProcess()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVPage>(this, "CreateSubProcess");
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="newShape">optional NetOffice.VisioApi.IVShape NewShape = 0</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop, object newShape)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop, newShape);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="page">NetOffice.VisioApi.IVPage page</param>
		/// <param name="objectToDrop">object objectToDrop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		public NetOffice.VisioApi.IVSelection MoveToSubprocess(NetOffice.VisioApi.IVPage page, object objectToDrop)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVSelection>(this, "MoveToSubprocess", page, objectToDrop);
		}

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="delFlags">Int32 delFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		public void DeleteEx(Int32 delFlags)
		{
			 Factory.ExecuteMethod(this, "DeleteEx", delFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		/// <param name="replaceFlags">optional Int32 ReplaceFlags = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop, object replaceFlags)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "ReplaceShape", masterOrMasterShortcutToDrop, replaceFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="masterOrMasterShortcutToDrop">object masterOrMasterShortcutToDrop</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 15, 16)]
		public NetOffice.VisioApi.IVShape ReplaceShape(object masterOrMasterShortcutToDrop)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "ReplaceShape", masterOrMasterShortcutToDrop);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="lineMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix</param>
		/// <param name="fillMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix</param>
		/// <param name="effectsMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix</param>
		/// <param name="fontMatrix">NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix</param>
		/// <param name="lineColor">NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor</param>
		/// <param name="fillColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor</param>
		/// <param name="shadowColor">NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor</param>
		/// <param name="fontColor">NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor</param>
		[SupportByVersion("Visio", 15, 16)]
		public void SetQuickStyle(NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices lineMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fillMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices effectsMatrix, NetOffice.VisioApi.Enums.VisQuickStyleMatrixIndices fontMatrix, NetOffice.VisioApi.Enums.VisQuickStyleColors lineColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fillColor, NetOffice.VisioApi.Enums.VisQuickStyleColors shadowColor, NetOffice.VisioApi.Enums.VisQuickStyleColors fontColor)
		{
			 Factory.ExecuteMethod(this, "SetQuickStyle", new object[]{ lineMatrix, fillMatrix, effectsMatrix, fontMatrix, lineColor, fillColor, shadowColor, fontColor });
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="changePictureFlags">optional Int32 ChangePictureFlags = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		public Double ChangePicture(string fileName, object changePictureFlags)
		{
			return Factory.ExecuteDoubleMethodGet(this, "ChangePicture", fileName, changePictureFlags);
		}

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		public Double ChangePicture(string fileName)
		{
			return Factory.ExecuteDoubleMethodGet(this, "ChangePicture", fileName);
		}

		#endregion

		#pragma warning restore
	}
}
