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
	/// Interface LPVISIOCELL 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class LPVISIOCELL : COMObject
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
                    _type = typeof(LPVISIOCELL);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOCELL(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCELL(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Error
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Error", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string Formula
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Formula", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Formula", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FormulaForce
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaForce", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaForce", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_Result(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "Result", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Result(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "Result", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_Result
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double Result(object unitsNameOrCode)
		{
			return get_Result(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Double get_ResultForce(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "ResultForce", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultForce(object unitsNameOrCode, Double value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "ResultForce", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultForce
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double ResultForce(object unitsNameOrCode)
		{
			return get_ResultForce(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double ResultIU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ResultIU", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ResultIU", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Double ResultIUForce
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ResultIUForce", paramsArray);
				return NetRuntimeSystem.Convert.ToDouble(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ResultIUForce", paramsArray);
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
		public Int16 Units
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Units", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
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
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string LocalName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LocalName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string RowName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowName", paramsArray);
			}
		}

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
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVShape Shape
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shape", paramsArray);
				NetOffice.VisioApi.IVShape newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVShape;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVStyle Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				NetOffice.VisioApi.IVStyle newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVStyle;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Section
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Section", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Row
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Row", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 Column
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Column", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 IsConstant
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsConstant", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 IsInherited
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsInherited", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		/// <param name="fRound">Int16 fRound</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ResultInt(object unitsNameOrCode, Int16 fRound)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode, fRound);
			object returnItem = Invoker.PropertyGet(this, "ResultInt", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultInt
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		/// <param name="fRound">Int16 fRound</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ResultInt(object unitsNameOrCode, Int16 fRound)
		{
			return get_ResultInt(unitsNameOrCode, fRound);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ResultFromInt(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "ResultFromInt", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultFromInt(object unitsNameOrCode, Int32 value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "ResultFromInt", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromInt
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ResultFromInt(object unitsNameOrCode)
		{
			return get_ResultFromInt(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_ResultFromIntForce(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "ResultFromIntForce", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ResultFromIntForce(object unitsNameOrCode, Int32 value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			Invoker.PropertySet(this, "ResultFromIntForce", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultFromIntForce
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 ResultFromIntForce(object unitsNameOrCode)
		{
			return get_ResultFromIntForce(unitsNameOrCode);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ResultStr(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "ResultStr", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ResultStr
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string ResultStr(object unitsNameOrCode)
		{
			return get_ResultStr(unitsNameOrCode);
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
		public NetOffice.VisioApi.IVRow ContainingRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ContainingRow", paramsArray);
				NetOffice.VisioApi.IVRow newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVRow;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FormulaU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaU", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FormulaForceU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FormulaForceU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "FormulaForceU", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string RowNameU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowNameU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RowNameU", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell InheritedValueSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InheritedValueSource", paramsArray);
				NetOffice.VisioApi.IVCell newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCell;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell InheritedFormulaSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InheritedFormulaSource", paramsArray);
				NetOffice.VisioApi.IVCell newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.VisioApi.IVCell;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell[] Dependents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dependents", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this,(object[])returnItem);
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
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public NetOffice.VisioApi.IVCell[] Precedents
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Precedents", paramsArray);
                ICOMObject[] newObject = Factory.CreateObjectArrayFromComProxy(this,(object[])returnItem);
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
		/// Get
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_ResultStrU(object unitsNameOrCode)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(unitsNameOrCode);
			object returnItem = Invoker.PropertyGet(this, "ResultStrU", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Alias for get_ResultStrU
		/// </summary>
		/// <param name="unitsNameOrCode">object UnitsNameOrCode</param>
		[SupportByVersionAttribute("Visio", 12,14,15,16)]
		public string ResultStrU(object unitsNameOrCode)
		{
			return get_ResultStrU(unitsNameOrCode);
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="cellObject">NetOffice.VisioApi.IVCell CellObject</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GlueTo(NetOffice.VisioApi.IVCell cellObject)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cellObject);
			Invoker.Method(this, "GlueTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="sheetObject">NetOffice.VisioApi.IVShape SheetObject</param>
		/// <param name="xPercent">Double xPercent</param>
		/// <param name="yPercent">Double yPercent</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void GlueToPos(NetOffice.VisioApi.IVShape sheetObject, Double xPercent, Double yPercent)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sheetObject, xPercent, yPercent);
			Invoker.Method(this, "GlueToPos", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void Trigger()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Trigger", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}