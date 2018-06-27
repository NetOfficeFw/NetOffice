using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVDispExtender 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVDispExtender : COMObject, NetOffice.VisioApi.IVDispExtender
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
                    _contractType = typeof(NetOffice.VisioApi.IVDispExtender);
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
                    _type = typeof(IVDispExtender);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVDispExtender() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object Object
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Object");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVDocument Document
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVDocument>(this, "Document");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object ShapeParent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ShapeParent");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVApplication Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVApplication>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVMaster Master
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "Master");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVCell get_Cells(string localeSpecificCellName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "Cells", typeof(NetOffice.VisioApi.IVCell), localeSpecificCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Cells
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Cells")]
		public virtual NetOffice.VisioApi.IVCell Cells(string localeSpecificCellName)
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
		public virtual NetOffice.VisioApi.IVCell get_CellsSRC(Int16 section, Int16 row, Int16 column)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsSRC", typeof(NetOffice.VisioApi.IVCell), section, row, column);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsSRC
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="column">Int16 column</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsSRC")]
		public virtual NetOffice.VisioApi.IVCell CellsSRC(Int16 section, Int16 row, Int16 column)
		{
			return get_CellsSRC(section, row, column);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Data1
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data1");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data1", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Data2
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data2");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data2", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Data3
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Data3");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Data3", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Help
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Help");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Help", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string NameID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_RowCount(Int16 section)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowCount", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowCount")]
		public virtual Int16 RowCount(Int16 section)
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
		public virtual Int16 get_RowsCellCount(Int16 section, Int16 row)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowsCellCount", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowsCellCount
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowsCellCount")]
		public virtual Int16 RowsCellCount(Int16 section, Int16 row)
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
		public virtual Int16 get_RowType(Int16 section, Int16 row)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowType", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_RowType(Int16 section, Int16 row, Int16 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "RowType", section, row, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowType
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowType")]
		public virtual Int16 RowType(Int16 section, Int16 row)
		{
			return get_RowType(section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVConnects Connects
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "Connects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ShapeIndex16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShapeIndex16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string StyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LineStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LineStyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineStyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FillStyle
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyle");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FillStyleKeepFmt
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillStyleKeepFmt");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillStyleKeepFmt", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_UniqueID(Int16 fUniqueID)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UniqueID", fUniqueID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_UniqueID
		/// </summary>
		/// <param name="fUniqueID">Int16 fUniqueID</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_UniqueID")]
		public virtual string UniqueID(Int16 fUniqueID)
		{
			return get_UniqueID(fUniqueID);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVPage ContainingPage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVPage>(this, "ContainingPage");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVMaster ContainingMaster
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVMaster>(this, "ContainingMaster");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape ContainingShape
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "ContainingShape");
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
		public virtual Int16 get_SectionExists(Int16 section, Int16 fExistsLocally)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "SectionExists", section, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SectionExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SectionExists")]
		public virtual Int16 SectionExists(Int16 section, Int16 fExistsLocally)
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
		public virtual Int16 get_RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RowExists", section, row, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RowExists
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RowExists")]
		public virtual Int16 RowExists(Int16 section, Int16 row, Int16 fExistsLocally)
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
		public virtual Int16 get_CellExists(string localeSpecificCellName, Int16 fExistsLocally)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellExists", localeSpecificCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExists
		/// </summary>
		/// <param name="localeSpecificCellName">string localeSpecificCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellExists")]
		public virtual Int16 CellExists(string localeSpecificCellName, Int16 fExistsLocally)
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
		public virtual Int16 get_CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellsSRCExists", section, row, column, fExistsLocally);
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
		public virtual Int16 CellsSRCExists(Int16 section, Int16 row, Int16 column, Int16 fExistsLocally)
		{
			return get_CellsSRCExists(section, row, column, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 LayerCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LayerCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVLayer get_Layer(Int16 index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVLayer>(this, "Layer", typeof(NetOffice.VisioApi.IVLayer), index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Layer
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Layer")]
		public virtual NetOffice.VisioApi.IVLayer Layer(Int16 index)
		{
			return get_Layer(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVEventList EventList
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVEventList>(this, "EventList");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 PersistsEvents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "PersistsEvents");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string ClassID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ClassID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), ProxyResult]
		public virtual object ShapeObject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ShapeObject");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ShapeID16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ShapeID16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVConnects FromConnects
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVConnects>(this, "FromConnects");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVHyperlink Hyperlink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVHyperlink>(this, "Hyperlink");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string ProgID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ProgID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectIsInherited
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectIsInherited");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ShapeID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ShapeID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ShapeIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ShapeIndex");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Index()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Index");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void VoidGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "VoidGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void BringForward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringForward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void BringToFront()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "BringToFront");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ConvertToGroup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertToGroup");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SendBackward()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendBackward");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SendToBack()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SendToBack");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ShapeCopy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShapeCopy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ShapeCut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShapeCut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void ShapeDelete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShapeDelete");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void VoidShapeDuplicate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "VoidShapeDuplicate");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AddSection(Int16 section)
		{
			return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddSection", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void DeleteSection(Int16 section)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteSection", section);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AddRow(Int16 section, Int16 row, Int16 rowTag)
		{
			return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddRow", section, row, rowTag);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void DeleteRow(Int16 section, Int16 row)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRow", section, row);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void SetCenter(Double xPos, Double yPos)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetCenter", xPos, yPos);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Export(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Export", fileName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="rowName">string rowName</param>
		/// <param name="rowTag">Int16 rowTag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AddNamedRow(Int16 section, string rowName, Int16 rowTag)
		{
			return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddNamedRow", section, rowName, rowTag);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="section">Int16 section</param>
		/// <param name="row">Int16 row</param>
		/// <param name="rowTag">Int16 rowTag</param>
		/// <param name="rowCount">Int16 rowCount</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 AddRows(Int16 section, Int16 row, Int16 rowTag, Int16 rowCount)
		{
			return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "AddRows", section, row, rowTag, rowCount);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVWindow OpenSheetWindow()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVWindow>(this, "OpenSheetWindow");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void GetFormulas(Int16[] sRCStream, out object[] formulaArray)
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
		public virtual void GetResults(Int16[] sRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray)
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
		public virtual Int16 SetFormulas(Int16[] sRCStream, object[] formulaArray, Int16 flags)
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
		public virtual Int16 SetResults(Int16[] sRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags)
		{
            object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)unitsNamesOrCodes, (object)resultArray, flags);
            object returnItem = Invoker.MethodReturn(this, "SetResults", paramsArray);
            return NetRuntimeSystem.Convert.ToInt16(returnItem);
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
		public virtual void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top)
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
		public virtual Int16 HitTest(Double xPos, Double yPos, Double tolerance)
		{
			return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "HitTest", xPos, yPos, tolerance);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape Group()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "Group");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVShape ShapeDuplicate()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.VisioApi.IVShape>(this, "ShapeDuplicate");
		}

		#endregion

		#pragma warning restore
	}
}


