using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Command15_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Command15_Deprecated : _ADO, NetOffice.ADODBApi.Command15_Deprecated
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
                    _contractType = typeof(NetOffice.ADODBApi.Command15_Deprecated);
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
                    _type = typeof(Command15_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Command15_Deprecated() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Connection_Deprecated ActiveConnection		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi._Connection_Deprecated>(this, "ActiveConnection", typeof(NetOffice.ADODBApi._Connection_Deprecated));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string CommandText
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 CommandTimeout
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CommandTimeout");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CommandTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual bool Prepared
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Prepared");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Prepared", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Parameters_Deprecated Parameters
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Parameters_Deprecated>(this, "Parameters", typeof(NetOffice.ADODBApi.Parameters_Deprecated));
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.CommandTypeEnum CommandType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CommandTypeEnum>(this, "CommandType");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CommandType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected, object parameters, object options)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", typeof(NetOffice.ADODBApi._Recordset_Deprecated), recordsAffected, parameters, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", typeof(NetOffice.ADODBApi._Recordset_Deprecated), recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Recordset_Deprecated Execute(object recordsAffected, object parameters)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Recordset_Deprecated>(this, "Execute", typeof(NetOffice.ADODBApi._Recordset_Deprecated), recordsAffected, parameters);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction, object size, object value)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated), new object[]{ name, type, direction, size, value });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated));
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated), name);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated), name, type);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated), name, type, direction);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi._Parameter_Deprecated CreateParameter(object name, object type, object direction, object size)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi._Parameter_Deprecated>(this, "CreateParameter", typeof(NetOffice.ADODBApi._Parameter_Deprecated), name, type, direction, size);
		}

		#endregion

		#pragma warning restore
	}
}


