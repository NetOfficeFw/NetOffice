using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.VisioApi
{
	///<summary>
	/// DispatchInterface IVCharacters 
	/// SupportByVersion Visio, 11,12,14,15,16
	///</summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IVCharacters : COMObject
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
                    _type = typeof(IVCharacters);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IVCharacters(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IVCharacters(string progId) : base(progId)
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 Begin
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Begin", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Begin", paramsArray);
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
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CharProps(Int16 cellIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cellIndex);
			object returnItem = Invoker.PropertyGet(this, "CharProps", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_CharProps(Int16 cellIndex, Int16 value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cellIndex);
			Invoker.PropertySet(this, "CharProps", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharProps
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CharProps(Int16 cellIndex)
		{
			return get_CharProps(cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CharPropsRow(Int16 biasLorR)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(biasLorR);
			object returnItem = Invoker.PropertyGet(this, "CharPropsRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 CharPropsRow(Int16 biasLorR)
		{
			return get_CharPropsRow(biasLorR);
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
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 End
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "End", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "End", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 FieldCategory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldCategory", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 FieldCode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldCode", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 FieldFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldFormat", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public string FieldFormula
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldFormula", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 IsField
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsField", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ParaProps(Int16 cellIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(cellIndex);
			object returnItem = Invoker.PropertyGet(this, "ParaProps", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ParaProps(Int16 cellIndex, Int16 value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cellIndex);
			Invoker.PropertySet(this, "ParaProps", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaProps
		/// </summary>
		/// <param name="cellIndex">Int16 CellIndex</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ParaProps(Int16 cellIndex)
		{
			return get_ParaProps(cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ParaPropsRow(Int16 biasLorR)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(biasLorR);
			object returnItem = Invoker.PropertyGet(this, "ParaPropsRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 ParaPropsRow(Int16 biasLorR)
		{
			return get_ParaPropsRow(biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_TabPropsRow(Int16 biasLorR)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(biasLorR);
			object returnItem = Invoker.PropertyGet(this, "TabPropsRow", paramsArray);
			return NetRuntimeSystem.Convert.ToInt16(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_TabPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 BiasLorR</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int16 TabPropsRow(Int16 biasLorR)
		{
			return get_TabPropsRow(biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 RunType</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_RunBegin(Int16 runType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(runType);
			object returnItem = Invoker.PropertyGet(this, "RunBegin", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunBegin
		/// </summary>
		/// <param name="runType">Int16 RunType</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 RunBegin(Int16 runType)
		{
			return get_RunBegin(runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 RunType</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_RunEnd(Int16 runType)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(runType);
			object returnItem = Invoker.PropertyGet(this, "RunEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunEnd
		/// </summary>
		/// <param name="runType">Int16 RunType</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public Int32 RunEnd(Int16 runType)
		{
			return get_RunEnd(runType);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string TextAsString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TextAsString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public object Text
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Text", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
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
		public string FieldFormulaU
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FieldFormulaU", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="formula">string Formula</param>
		/// <param name="format">Int16 Format</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddCustomField(string formula, Int16 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, format);
			Invoker.Method(this, "AddCustomField", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="category">Int16 Category</param>
		/// <param name="code">Int16 Code</param>
		/// <param name="format">Int16 Format</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddField(Int16 category, Int16 code, Int16 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category, code, format);
			Invoker.Method(this, "AddField", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
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
		/// <param name="formula">string Formula</param>
		/// <param name="format">Int16 Format</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddCustomFieldU(string formula, Int16 format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(formula, format);
			Invoker.Method(this, "AddCustomFieldU", paramsArray);
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
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories Category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes Code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats Format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		/// <param name="calendarID">optional Int32 CalendarID = -1</param>
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID, object calendarID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category, code, format, langID, calendarID);
			Invoker.Method(this, "AddFieldEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories Category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes Code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category, code, format);
			Invoker.Method(this, "AddFieldEx", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories Category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes Code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats Format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(category, code, format, langID);
			Invoker.Method(this, "AddFieldEx", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}