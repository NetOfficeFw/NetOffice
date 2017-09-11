using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Recordset 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Recordset : Recordset21
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
                    _type = typeof(_Recordset);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Recordset(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Recordset(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Recordset(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Properties>(this, "Properties", NetOffice.ADODBApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.PositionEnum AbsolutePosition
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.PositionEnum>(this, "AbsolutePosition");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AbsolutePosition", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public object ActiveConnection
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveConnection");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public bool BOF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BOF");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object Bookmark
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Bookmark");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Bookmark", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 CacheSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CacheSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CacheSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorTypeEnum>(this, "CursorType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CursorType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public bool EOF
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EOF");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Fields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Fields>(this, "Fields", NetOffice.ADODBApi.Fields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.LockTypeEnum LockType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.LockTypeEnum>(this, "LockType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LockType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 MaxRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 RecordCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RecordCount");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5), ProxyResult]
		public object Source
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Source");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Source", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.PositionEnum AbsolutePage
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.PositionEnum>(this, "AbsolutePage");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AbsolutePage", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.EditModeEnum EditMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.EditModeEnum>(this, "EditMode");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object Filter
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Filter");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Filter", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 PageCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageCount");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 PageSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string Sort
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Sort");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Sort", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 Status
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Status");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 State
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.CursorLocationEnum CursorLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CursorLocationEnum>(this, "CursorLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CursorLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Collect(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Collect", index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Collect(object index, object value)
		{
			Factory.ExecutePropertySet(this, "Collect", index, value);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.1,2.5), Redirect("get_Collect")]
		public object Collect(object index)
		{
			return get_Collect(index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.MarshalOptionsEnum MarshalOptions
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.MarshalOptionsEnum>(this, "MarshalOptions");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarshalOptions", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public string Index
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Index");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
		public object DataSource
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DataSource");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.5), ProxyResult]
		public object ActiveCommand
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ActiveCommand");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public bool StayInSync
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StayInSync");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StayInSync", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string DataMember
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DataMember");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataMember", value);
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
		public void AddNew(object fieldList, object values)
		{
			 Factory.ExecuteMethod(this, "AddNew", fieldList, values);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void AddNew()
		{
			 Factory.ExecuteMethod(this, "AddNew");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fieldList">optional object fieldList</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void AddNew(object fieldList)
		{
			 Factory.ExecuteMethod(this, "AddNew", fieldList);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void CancelUpdate()
		{
			 Factory.ExecuteMethod(this, "CancelUpdate");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Delete(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "Delete", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		/// <param name="fields">optional object fields</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object GetRows(object rows, object start, object fields)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows", rows, start, fields);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object GetRows()
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object GetRows(object rows)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows", rows);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="rows">optional Int32 Rows = -1</param>
		/// <param name="start">optional object start</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public object GetRows(object rows, object start)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetRows", rows, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Move(Int32 numRecords, object start)
		{
			 Factory.ExecuteMethod(this, "Move", numRecords, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="numRecords">Int32 numRecords</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Move(Int32 numRecords)
		{
			 Factory.ExecuteMethod(this, "Move", numRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void MoveNext()
		{
			 Factory.ExecuteMethod(this, "MoveNext");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void MovePrevious()
		{
			 Factory.ExecuteMethod(this, "MovePrevious");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void MoveFirst()
		{
			 Factory.ExecuteMethod(this, "MoveFirst");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void MoveLast()
		{
			 Factory.ExecuteMethod(this, "MoveLast");
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
		public void Open(object source, object activeConnection, object cursorType, object lockType, object options)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, cursorType, lockType, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Open(object source)
		{
			 Factory.ExecuteMethod(this, "Open", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Open(object source, object activeConnection)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="cursorType">optional NetOffice.ADODBApi.Enums.CursorTypeEnum CursorType = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Open(object source, object activeConnection, object cursorType)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection, cursorType);
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
		public void Open(object source, object activeConnection, object cursorType, object lockType)
		{
			 Factory.ExecuteMethod(this, "Open", source, activeConnection, cursorType, lockType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Requery(object options)
		{
			 Factory.ExecuteMethod(this, "Requery", options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Requery()
		{
			 Factory.ExecuteMethod(this, "Requery");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void _xResync(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "_xResync", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void _xResync()
		{
			 Factory.ExecuteMethod(this, "_xResync");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		/// <param name="values">optional object values</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Update(object fields, object values)
		{
			 Factory.ExecuteMethod(this, "Update", fields, values);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Update()
		{
			 Factory.ExecuteMethod(this, "Update");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="fields">optional object fields</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Update(object fields)
		{
			 Factory.ExecuteMethod(this, "Update", fields);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset _xClone()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "_xClone");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void UpdateBatch(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "UpdateBatch", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void UpdateBatch()
		{
			 Factory.ExecuteMethod(this, "UpdateBatch");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void CancelBatch(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "CancelBatch", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void CancelBatch()
		{
			 Factory.ExecuteMethod(this, "CancelBatch");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">optional object recordsAffected</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Recordset NextRecordset(object recordsAffected)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "NextRecordset", recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset NextRecordset()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "NextRecordset");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="cursorOptions">NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public bool Supports(NetOffice.ADODBApi.Enums.CursorOptionEnum cursorOptions)
		{
			return Factory.ExecuteBoolMethodGet(this, "Supports", cursorOptions);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		/// <param name="start">optional object start</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Find(string criteria, object skipRecords, object searchDirection, object start)
		{
			 Factory.ExecuteMethod(this, "Find", criteria, skipRecords, searchDirection, start);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Find(string criteria)
		{
			 Factory.ExecuteMethod(this, "Find", criteria);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Find(string criteria, object skipRecords)
		{
			 Factory.ExecuteMethod(this, "Find", criteria, skipRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="criteria">string criteria</param>
		/// <param name="skipRecords">optional Int32 SkipRecords = 0</param>
		/// <param name="searchDirection">optional NetOffice.ADODBApi.Enums.SearchDirectionEnum SearchDirection = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public void Find(string criteria, object skipRecords, object searchDirection)
		{
			 Factory.ExecuteMethod(this, "Find", criteria, skipRecords, searchDirection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		/// <param name="seekOption">optional NetOffice.ADODBApi.Enums.SeekEnum SeekOption = 1</param>
		[SupportByVersion("ADODB", 2.1)]
		public void Seek(object keyValues, object seekOption)
		{
			 Factory.ExecuteMethod(this, "Seek", keyValues, seekOption);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.1)]
		public void Seek(object keyValues)
		{
			 Factory.ExecuteMethod(this, "Seek", keyValues);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}
        
		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		/// <param name="nullExpr">optional string nullExpr</param>
		[SupportByVersion("ADODB", 2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter, object nullExpr)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", new object[]{ stringFormat, numRows, columnDelimeter, rowDelimeter, nullExpr });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string GetString()
		{
			return Factory.ExecuteStringMethodGet(this, "GetString");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string GetString(object stringFormat)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string GetString(object stringFormat, object numRows)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="stringFormat">optional NetOffice.ADODBApi.Enums.StringFormatEnum StringFormat = 2</param>
		/// <param name="numRows">optional Int32 NumRows = -1</param>
		/// <param name="columnDelimeter">optional string columnDelimeter</param>
		/// <param name="rowDelimeter">optional string rowDelimeter</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string GetString(object stringFormat, object numRows, object columnDelimeter, object rowDelimeter)
		{
			return Factory.ExecuteStringMethodGet(this, "GetString", stringFormat, numRows, columnDelimeter, rowDelimeter);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="bookmark1">object bookmark1</param>
		/// <param name="bookmark2">object bookmark2</param>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.CompareEnum CompareBookmarks(object bookmark1, object bookmark2)
		{
			return Factory.ExecuteEnumMethodGet<NetOffice.ADODBApi.Enums.CompareEnum>(this, "CompareBookmarks", bookmark1, bookmark2);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="lockType">optional NetOffice.ADODBApi.Enums.LockTypeEnum LockType = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Recordset Clone(object lockType)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone", lockType);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi._Recordset Clone()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Clone");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		/// <param name="resyncValues">optional NetOffice.ADODBApi.Enums.ResyncEnum ResyncValues = 2</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Resync(object affectRecords, object resyncValues)
		{
			 Factory.ExecuteMethod(this, "Resync", affectRecords, resyncValues);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Resync()
		{
			 Factory.ExecuteMethod(this, "Resync");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="affectRecords">optional NetOffice.ADODBApi.Enums.AffectEnum AffectRecords = 3</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Resync(object affectRecords)
		{
			 Factory.ExecuteMethod(this, "Resync", affectRecords);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		/// <param name="persistFormat">optional NetOffice.ADODBApi.Enums.PersistFormatEnum PersistFormat = 0</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Save(object destination, object persistFormat)
		{
			 Factory.ExecuteMethod(this, "Save", destination, persistFormat);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destination">optional object destination</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Save(object destination)
		{
			 Factory.ExecuteMethod(this, "Save", destination);
		}

		#endregion

		#pragma warning restore
	}
}
