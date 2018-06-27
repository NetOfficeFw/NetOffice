using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.VisioApi;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// Interface LPVISIOCHARS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class LPVISIOCHARS : COMObject, NetOffice.VisioApi.LPVISIOCHARS
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
                    _contractType = typeof(NetOffice.VisioApi.LPVISIOCHARS);
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
                    _type = typeof(LPVISIOCHARS);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LPVISIOCHARS() : base()
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 Begin
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Begin");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Begin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int32 CharCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CharCount");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_CharProps(Int16 cellIndex)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CharProps", cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_CharProps(Int16 cellIndex, Int16 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "CharProps", cellIndex, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharProps")]
		public virtual Int16 CharProps(Int16 cellIndex)
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
		public virtual Int16 get_CharPropsRow(Int16 biasLorR)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "CharPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_CharPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_CharPropsRow")]
		public virtual Int16 CharPropsRow(Int16 biasLorR)
		{
			return get_CharPropsRow(biasLorR);
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
		public virtual Int32 End
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "End");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "End", value);
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 FieldCategory
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FieldCategory");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 FieldCode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FieldCode");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 FieldFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "FieldFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual string FieldFormula
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FieldFormula");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual Int16 IsField
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "IsField");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_ParaProps(Int16 cellIndex)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ParaProps", cellIndex);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_ParaProps(Int16 cellIndex, Int16 value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "ParaProps", cellIndex, value);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaProps
		/// </summary>
		/// <param name="cellIndex">Int16 cellIndex</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaProps")]
		public virtual Int16 ParaProps(Int16 cellIndex)
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
		public virtual Int16 get_ParaPropsRow(Int16 biasLorR)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ParaPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_ParaPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_ParaPropsRow")]
		public virtual Int16 ParaPropsRow(Int16 biasLorR)
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
		public virtual Int16 get_TabPropsRow(Int16 biasLorR)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TabPropsRow", biasLorR);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_TabPropsRow
		/// </summary>
		/// <param name="biasLorR">Int16 biasLorR</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_TabPropsRow")]
		public virtual Int16 TabPropsRow(Int16 biasLorR)
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
		public virtual Int32 get_RunBegin(Int16 runType)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RunBegin", runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunBegin
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunBegin")]
		public virtual Int32 RunBegin(Int16 runType)
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
		public virtual Int32 get_RunEnd(Int16 runType)
		{
			return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RunEnd", runType);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_RunEnd
		/// </summary>
		/// <param name="runType">Int16 runType</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_RunEnd")]
		public virtual Int32 RunEnd(Int16 runType)
		{
			return get_RunEnd(runType);
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
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual string TextAsString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "TextAsString");
			}
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual object Text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Text", value);
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
		public virtual string FieldFormulaU
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FieldFormulaU");
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddCustomField(string formula, Int16 format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddCustomField", formula, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">Int16 category</param>
		/// <param name="code">Int16 code</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddField(Int16 category, Int16 code, Int16 format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddField", category, code, format);
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Copy()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Cut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void Paste()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="formula">string formula</param>
		/// <param name="format">Int16 format</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddCustomFieldU(string formula, Int16 format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddCustomFieldU", formula, format);
		}

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
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		/// <param name="langID">optional Int32 LangID = 0</param>
		/// <param name="calendarID">optional Int32 CalendarID = -1</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID, object calendarID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddFieldEx", new object[]{ category, code, format, langID, calendarID });
		}

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="category">NetOffice.VisioApi.Enums.VisFieldCategories category</param>
		/// <param name="code">NetOffice.VisioApi.Enums.VisFieldCodes code</param>
		/// <param name="format">NetOffice.VisioApi.Enums.VisFieldFormats format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddFieldEx", category, code, format);
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
		public virtual void AddFieldEx(NetOffice.VisioApi.Enums.VisFieldCategories category, NetOffice.VisioApi.Enums.VisFieldCodes code, NetOffice.VisioApi.Enums.VisFieldFormats format, object langID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddFieldEx", category, code, format, langID);
		}

		#endregion

		#pragma warning restore
	}
}

