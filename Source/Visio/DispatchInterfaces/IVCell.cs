using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVCell 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IVCell : COMObject
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
                    _type = typeof(IVCell);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IVCell(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVCell(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCell(string progId) : base(progId)
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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Error
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Error");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Formula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Formula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Formula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormulaForce
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaForce");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaForce", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_Result(object unitsNameOrCode)
		{
			return Factory.ExecuteDoublePropertyGet(this, "Result", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Result(object unitsNameOrCode, Double value)
		{
			Factory.ExecutePropertySet(this, "Result", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Result
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_Result")]
		public Double Result(object unitsNameOrCode)
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
		public Double get_ResultForce(object unitsNameOrCode)
		{
			return Factory.ExecuteDoublePropertyGet(this, "ResultForce", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultForce(object unitsNameOrCode, Double value)
		{
			Factory.ExecutePropertySet(this, "ResultForce", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultForce
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultForce")]
		public Double ResultForce(object unitsNameOrCode)
		{
			return get_ResultForce(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double ResultIU
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "ResultIU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResultIU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Double ResultIUForce
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "ResultIUForce");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResultIUForce", value);
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
		public Int16 Units
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Units");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string LocalName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LocalName");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string RowName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RowName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowName", value);
			}
		}

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
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVShape>(this, "Shape");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVStyle Style
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVStyle>(this, "Style");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Section
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Section");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Row
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Row");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 Column
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "Column");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IsConstant
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsConstant");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IsInherited
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsInherited");
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
		public Int32 get_ResultInt(object unitsNameOrCode, Int16 fRound)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ResultInt", unitsNameOrCode, fRound);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultInt
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		/// <param name="fRound">Int16 fRound</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultInt")]
		public Int32 ResultInt(object unitsNameOrCode, Int16 fRound)
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
		public Int32 get_ResultFromInt(object unitsNameOrCode)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ResultFromInt", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultFromInt(object unitsNameOrCode, Int32 value)
		{
			Factory.ExecutePropertySet(this, "ResultFromInt", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromInt
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultFromInt")]
		public Int32 ResultFromInt(object unitsNameOrCode)
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
		public Int32 get_ResultFromIntForce(object unitsNameOrCode)
		{
			return Factory.ExecuteInt32PropertyGet(this, "ResultFromIntForce", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultFromIntForce(object unitsNameOrCode, Int32 value)
		{
			Factory.ExecutePropertySet(this, "ResultFromIntForce", unitsNameOrCode, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromIntForce
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultFromIntForce")]
		public Int32 ResultFromIntForce(object unitsNameOrCode)
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
		public string get_ResultStr(object unitsNameOrCode)
		{
			return Factory.ExecuteStringPropertyGet(this, "ResultStr", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultStr
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ResultStr")]
		public string ResultStr(object unitsNameOrCode)
		{
			return get_ResultStr(unitsNameOrCode);
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
		[BaseResult]
		public NetOffice.VisioApi.IVRow ContainingRow
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVRow>(this, "ContainingRow");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormulaU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FormulaForceU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FormulaForceU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FormulaForceU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string RowNameU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RowNameU");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RowNameU", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVCell InheritedValueSource
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "InheritedValueSource");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.VisioApi.IVCell InheritedFormulaSource
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VisioApi.IVCell>(this, "InheritedFormulaSource");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell[] Dependents
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
		public NetOffice.VisioApi.IVCell[] Precedents
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
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ResultStrU(object unitsNameOrCode)
		{
			return Factory.ExecuteStringPropertyGet(this, "ResultStrU", unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ResultStrU
		/// </summary>
		/// <param name="unitsNameOrCode">object unitsNameOrCode</param>
		[SupportByVersion("Visio", 12,14,15,16), Redirect("get_ResultStrU")]
		public string ResultStrU(object unitsNameOrCode)
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
		public void GlueTo(NetOffice.VisioApi.IVCell cellObject)
		{
			 Factory.ExecuteMethod(this, "GlueTo", cellObject);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape sheetObject</param>
		/// <param name="xPercent">Double xPercent</param>
		/// <param name="yPercent">Double yPercent</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void GlueToPos(NetOffice.VisioApi.IVShape sheetObject, Double xPercent, Double yPercent)
		{
			 Factory.ExecuteMethod(this, "GlueToPos", sheetObject, xPercent, yPercent);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Trigger()
		{
			 Factory.ExecuteMethod(this, "Trigger");
		}

		#endregion

		#pragma warning restore
	}
}
