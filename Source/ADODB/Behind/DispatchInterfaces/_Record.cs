using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface _Record 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Record : _ADO, NetOffice.ADODBApi._Record
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
                    _contractType = typeof(NetOffice.ADODBApi._Record);
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
                    _type = typeof(_Record);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Record() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual object ActiveConnection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ActiveConnection");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.ObjectStateEnum State
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ObjectStateEnum>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual object Source
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Source");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Source", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ConnectModeEnum>(this, "Mode");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Mode", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string ParentURL
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ParentURL");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Fields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Fields>(this, "Fields", typeof(NetOffice.ADODBApi.Fields));
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.RecordTypeEnum RecordType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.RecordTypeEnum>(this, "RecordType");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source, object destination, object userName, object password, object options, object async)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", new object[]{ source, destination, userName, password, options, async });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source, object destination)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", source, destination);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source, object destination, object userName)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", source, destination, userName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source, object destination, object userName, object password)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", source, destination, userName, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.MoveRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string MoveRecord(object source, object destination, object userName, object password, object options)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "MoveRecord", new object[]{ source, destination, userName, password, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source, object destination, object userName, object password, object options, object async)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", new object[]{ source, destination, userName, password, options, async });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source, object destination)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", source, destination);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source, object destination, object userName)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", source, destination, userName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source, object destination, object userName, object password)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", source, destination, userName, password);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string Source = </param>
		/// <param name="destination">optional string Destination = </param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.CopyRecordOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CopyRecord(object source, object destination, object userName, object password, object options)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CopyRecord", new object[]{ source, destination, userName, password, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string source</param>
		/// <param name="async">optional bool async</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void DeleteRecord(object source, object async)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRecord", source, async);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void DeleteRecord()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRecord");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional string source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void DeleteRecord(object source)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRecord", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options, userName, password });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection, object mode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection, mode);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection, object mode, object createOptions)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, activeConnection, mode, createOptions);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection, object mode, object createOptions, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="activeConnection">optional object activeConnection</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum mode</param>
		/// <param name="createOptions">optional NetOffice.ADODBApi.Enums.RecordCreateOptionsEnum CreateOptions = -1</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.RecordOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object activeConnection, object mode, object createOptions, object options, object userName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[]{ source, activeConnection, mode, createOptions, options, userName });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		[BaseResult]
		public virtual NetOffice.ADODBApi._Recordset GetChildren()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "GetChildren");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Cancel()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Cancel");
		}

		#endregion

		#pragma warning restore
	}
}


