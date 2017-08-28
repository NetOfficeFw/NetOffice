using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOCHARS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOCHARS : COMObject
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
                    _type = typeof(LPVISIOCHARS);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public LPVISIOCHARS(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public LPVISIOCHARS(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public LPVISIOCHARS(string progId) : base(progId)
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 Begin
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Begin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Begin", value);
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
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CharProps(Int16 cellIndex)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CharProps", cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_CharProps(Int16 cellIndex, Int16 value)
		{
			Factory.ExecutePropertySet(this, "CharProps", cellIndex, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharProps")]
		public Int16 CharProps(Int16 cellIndex)
		{
			return get_CharProps(cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_CharPropsRow(Int16 biasLorR)
		{
			return Factory.ExecuteInt16PropertyGet(this, "CharPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharPropsRow")]
		public Int16 CharPropsRow(Int16 biasLorR)
		{
			return get_CharPropsRow(biasLorR);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int32 End
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "End");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "End", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 FieldCategory
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FieldCategory");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 FieldCode
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FieldCode");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 FieldFormat
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FieldFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public string FieldFormula
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FieldFormula");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public Int16 IsField
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "IsField");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ParaProps(Int16 cellIndex)
		{
			return Factory.ExecuteInt16PropertyGet(this, "ParaProps", cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ParaProps(Int16 cellIndex, Int16 value)
		{
			Factory.ExecutePropertySet(this, "ParaProps", cellIndex, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaProps")]
		public Int16 ParaProps(Int16 cellIndex)
		{
			return get_ParaProps(cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ParaPropsRow(Int16 biasLorR)
		{
			return Factory.ExecuteInt16PropertyGet(this, "ParaPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaPropsRow")]
		public Int16 ParaPropsRow(Int16 biasLorR)
		{
			return get_ParaPropsRow(biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_TabPropsRow(Int16 biasLorR)
		{
			return Factory.ExecuteInt16PropertyGet(this, "TabPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_TabPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_TabPropsRow")]
		public Int16 TabPropsRow(Int16 biasLorR)
		{
			return get_TabPropsRow(biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_RunBegin(Int16 runType)
		{
			return Factory.ExecuteInt32PropertyGet(this, "RunBegin", runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunBegin
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunBegin")]
		public Int32 RunBegin(Int16 runType)
		{
			return get_RunBegin(runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 get_RunEnd(Int16 runType)
		{
			return Factory.ExecuteInt32PropertyGet(this, "RunEnd", runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunEnd
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunEnd")]
		public Int32 RunEnd(Int16 runType)
		{
			return get_RunEnd(runType);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string TextAsString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextAsString");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public object Text
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Text", value);
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
		public string FieldFormulaU
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FieldFormulaU");
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddCustomField(string formula, Int16 format)
		{
			 Factory.ExecuteMethod(this, "AddCustomField", formula, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">Int16 category</param>
		/// <param name="code">Int16 code</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddField(Int16 category, Int16 code, Int16 format)
		{
			 Factory.ExecuteMethod(this, "AddField", category, code, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddCustomFieldU(string formula, Int16 format)
		{
			 Factory.ExecuteMethod(this, "AddCustomFieldU", formula, format);
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
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		/// <param name="calendarID">optional Int32 CalendarID = -1</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID, object calendarID)
		{
			 Factory.ExecuteMethod(this, "AddFieldEx", new object[]{ category, code, format, langID, calendarID });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format)
		{
			 Factory.ExecuteMethod(this, "AddFieldEx", category, code, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID)
		{
			 Factory.ExecuteMethod(this, "AddFieldEx", category, code, format, langID);
		}

		#endregion

		#pragma warning restore
	}
}
