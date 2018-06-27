using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOCELL 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOCELL : COMObject, NetOffice.VisioApi.LPVISIOCELL
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOCELL);
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
                    _type = typeof(LPVISIOCELL);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOCELL() : base()
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
		public virtual Int16 ObjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ObjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Error
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Error");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Formula
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FormulaForce
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaForce");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaForce", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double get_Result(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Result", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_Result(object unitsNameOrCode, Double value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "Result", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Result
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Result")]
		public virtual Double Result(object unitsNameOrCode)
		{
			return get_Result(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Double get_ResultForce(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "ResultForce", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_ResultForce(object unitsNameOrCode, Double value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "ResultForce", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultForce
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultForce")]
		public virtual Double ResultForce(object unitsNameOrCode)
		{
			return get_ResultForce(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Double ResultIU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "ResultIU");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResultIU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Double ResultIUForce
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "ResultIUForce");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ResultIUForce", value);
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
		public virtual Int16 Units
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Units");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string LocalName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "LocalName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string RowName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RowName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowName", value);
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
		public virtual NetOffice.VisioApi.IVStyle Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "Style");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Section
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Section");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Row
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Row");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 Column
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Column");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IsConstant
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IsConstant");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IsInherited
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IsInherited");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		/// <param name="fRound">Int16 fRound</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ResultInt(object unitsNameOrCode, Int16 fRound)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ResultInt", unitsNameOrCode, fRound);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultInt
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		/// <param name="fRound">Int16 fRound</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultInt")]
		public virtual Int32 ResultInt(object unitsNameOrCode, Int16 fRound)
		{
			return get_ResultInt(unitsNameOrCode, fRound);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ResultFromInt(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ResultFromInt", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_ResultFromInt(object unitsNameOrCode, Int32 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "ResultFromInt", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromInt
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultFromInt")]
		public virtual Int32 ResultFromInt(object unitsNameOrCode)
		{
			return get_ResultFromInt(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 get_ResultFromIntForce(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ResultFromIntForce", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_ResultFromIntForce(object unitsNameOrCode, Int32 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "ResultFromIntForce", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromIntForce
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultFromIntForce")]
		public virtual Int32 ResultFromIntForce(object unitsNameOrCode)
		{
			return get_ResultFromIntForce(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_ResultStr(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResultStr", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultStr
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultStr")]
		public virtual string ResultStr(object unitsNameOrCode)
		{
			return get_ResultStr(unitsNameOrCode);
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
		[BaseResult]
		public virtual NetOffice.VisioApi.IVRow ContainingRow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVRow>(this, "ContainingRow");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FormulaU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaU");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FormulaForceU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaForceU");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaForceU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string RowNameU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RowNameU");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RowNameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVCell InheritedValueSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "InheritedValueSource");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.VisioApi.IVCell InheritedFormulaSource
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "InheritedFormulaSource");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVCell[] Dependents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dependents", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this,(object[])returnItem, false);
				NetOffice.VisioApi.IVCell[] returnArray = new NetOffice.VisioApi.IVCell[newObject.Length];
				for (int i = 0; i < newObject.Length; i++)
					returnArray[i] = newObject[i] as NetOffice.VisioApi.IVCell;
				return returnArray;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual NetOffice.VisioApi.IVCell[] Precedents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Precedents", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this,(object[])returnItem, false);
				NetOffice.VisioApi.IVCell[] returnArray = new NetOffice.VisioApi.IVCell[newObject.Length];
				for (int i = 0; i < newObject.Length; i++)
					returnArray[i] = newObject[i] as NetOffice.VisioApi.IVCell;
				return returnArray;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ContainingPageID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingPageID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 ContainingMasterID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContainingMasterID");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string get_ResultStrU(object unitsNameOrCode)
		{
			return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ResultStrU", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ResultStrU
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 12,14,15,16), Redirect("get_ResultStrU")]
		public virtual string ResultStrU(object unitsNameOrCode)
		{
			return get_ResultStrU(unitsNameOrCode);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cellObject">NetOffice.VisioApi.IVCell cellObject</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void GlueTo(NetOffice.VisioApi.IVCell cellObject)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GlueTo", cellObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="xPercent">Double xPercent</param>
		/// <param name="yPercent">Double yPercent</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void GlueToPos(NetOffice.VisioApi.IVShape sheetObject, Double xPercent, Double yPercent)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "GlueToPos", sheetObject, xPercent, yPercent);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Trigger()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Trigger");
		}

		#endregion

		#pragma warning restore
	}
}

