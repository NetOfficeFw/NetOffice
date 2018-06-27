using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Recordset15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset15 : _ADO, NetOffice.ADODBApi.Recordset15
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
                    _contractType = typeof(NetOffice.ADODBApi.Recordset15);
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
                    _type = typeof(Recordset15);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Recordset15() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.PositionEnum AbsolutePosition
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.PositionEnum>(this, "AbsolutePosition");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AbsolutePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public virtual object ActiveConnection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ActiveConnection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual bool BOF
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "BOF");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object Bookmark
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Bookmark");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 CacheSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CacheSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CacheSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorTypeEnum>(this, "CursorType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CursorType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual bool EOF
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EOF");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Fields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Fields>(this, "Fields", typeof(NetOffice.ADODBApi.Fields));
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.LockTypeEnum LockType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.LockTypeEnum>(this, "LockType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LockType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 MaxRecords
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 RecordCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public virtual object Source
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Source");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Source", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.PositionEnum AbsolutePage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.PositionEnum>(this, "AbsolutePage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AbsolutePage", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.EditModeEnum EditMode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.EditModeEnum>(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object Filter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Filter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 PageCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageCount");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 PageSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual string Sort
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Sort");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Sort", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 Status
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Status");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual Int32 State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorLocationEnum>(this, "CursorLocation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CursorLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_Collect(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Collect", index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_Collect(object index, object value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "Collect", index, value);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5), Redirect("get_Collect")]
		public virtual object Collect(object index)
		{
			return get_Collect(index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi.Enums.MarshalOptionsEnum MarshalOptions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.MarshalOptionsEnum>(this, "MarshalOptions");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarshalOptions", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void AddNew(object fieldList, object values)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddNew", fieldList, values);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void AddNew()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddNew");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void AddNew(object fieldList)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddNew", fieldList);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void CancelUpdate()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Delete(object affectRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		/// <param name="fields">optional object fields</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object GetRows(object rows, object start, object fields)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows", rows, start, fields);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object GetRows()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object GetRows(object rows)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows", rows);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual object GetRows(object rows, object start)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetRows", rows, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Move(Int32 numRecords, object start)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", numRecords, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Move(Int32 numRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Move", numRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void MoveNext()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveNext");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void MovePrevious()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MovePrevious");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void MoveFirst()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveFirst");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void MoveLast()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "MoveLast");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open(object source, object activeConnection, object cursorType, object lockType, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, cursorType, lockType, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open(object source)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open(object source, object activeConnection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open(object source, object activeConnection, object cursorType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection, cursorType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Open(object source, object activeConnection, object cursorType, object lockType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection, cursorType, lockType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Requery(object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery", options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Requery()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void _xResync(object affectRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_xResync", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void _xResync()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_xResync");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Update(object fields, object values)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", fields, values);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Update()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Update(object fields)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Update", fields);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi._Recordset _xClone()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "_xClone");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void UpdateBatch(object affectRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateBatch", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void UpdateBatch()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateBatch");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void CancelBatch(object affectRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelBatch", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void CancelBatch()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CancelBatch");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">optional object recordsAffected</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public virtual NetOffice.ADODBApi._Recordset NextRecordset(object recordsAffected)
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "NextRecordset", recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual NetOffice.ADODBApi._Recordset NextRecordset()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "NextRecordset");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="cursorOptions">NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual bool Supports(NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Supports", cursorOptions);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Find(string criteria, object skipRecords, object searchDirection, object start)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Find", criteria, skipRecords, searchDirection, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Find(string criteria)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Find", criteria);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Find(string criteria, object skipRecords)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Find", criteria, skipRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public virtual void Find(string criteria, object skipRecords, object searchDirection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Find", criteria, skipRecords, searchDirection);
		}

		#endregion

		#pragma warning restore
	}
}


