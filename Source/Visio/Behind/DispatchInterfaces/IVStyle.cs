using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// DispatchInterface IVStyle 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVStyle : COMObject, NetOffice.VisioApi.IVStyle
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
                    _contractType = typeof(NetOffice.VisioApi.IVStyle);
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
                    _type = typeof(IVStyle);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IVStyle() : base()
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
		public virtual Int16 Stat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Stat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

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
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 Index16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index16");
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string BasedOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BasedOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BasedOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string TextBasedOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextBasedOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TextBasedOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LineBasedOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LineBasedOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineBasedOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FillBasedOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FillBasedOn");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FillBasedOn", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IncludesText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IncludesText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IncludesText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IncludesLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IncludesLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IncludesLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IncludesFill
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IncludesFill");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IncludesFill", value);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 ID16
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ID16");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVSection get_Section(Int16 index)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVSection>(this, "Section", typeof(NetOffice.VisioApi.IVSection), index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Section
		/// </summary>
		/// <param name="index">Int16 index</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Section")]
		public virtual NetOffice.VisioApi.IVSection Section(Int16 index)
		{
			return get_Section(index);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Hidden
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Hidden");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string NameU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameU");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual NetOffice.VisioApi.IVCell get_CellsU(string localeIndependentCellName)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "CellsU", typeof(NetOffice.VisioApi.IVCell), localeIndependentCellName);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellsU")]
		public virtual NetOffice.VisioApi.IVCell CellsU(string localeIndependentCellName)
		{
			return get_CellsU(localeIndependentCellName);
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
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CellExistsU", localeIndependentCellName, fExistsLocally);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CellExistsU
		/// </summary>
		/// <param name="localeIndependentCellName">string localeIndependentCellName</param>
		/// <param name="fExistsLocally">Int16 fExistsLocally</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CellExistsU")]
		public virtual Int16 CellExistsU(string localeIndependentCellName, Int16 fExistsLocally)
		{
			return get_CellExistsU(localeIndependentCellName, fExistsLocally);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
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
		/// <param name="sRCStream">Int16[] sRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void GetFormulasU(Int16[] sRCStream, out object[] formulaArray)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			formulaArray = null;
			object[] paramsArray = Invoker.ValidateParamsArray((object)sRCStream, (object)formulaArray);
			Invoker.Method(this, "GetFormulasU", paramsArray, modifiers);
			formulaArray = (object[])paramsArray[1];
		}

		#endregion

		#pragma warning restore
	}
}


