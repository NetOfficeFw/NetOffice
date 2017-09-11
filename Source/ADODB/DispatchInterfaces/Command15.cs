using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Command15 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Command15 : _ADO
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
                    _type = typeof(Command15);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Command15(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Command15(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Command15(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Connection ActiveConnection
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.ADODBApi._Connection>(this, "ActiveConnection");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ActiveConnection", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string CommandText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CommandText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandText", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public Int32 CommandTimeout
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CommandTimeout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CommandTimeout", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public bool Prepared
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Prepared");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Prepared", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Parameters Parameters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Parameters>(this, "Parameters", NetOffice.ADODBApi.Parameters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.CommandTypeEnum CommandType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.CommandTypeEnum>(this, "CommandType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CommandType", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1,2.5)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected, object parameters, object options)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Execute", recordsAffected, parameters, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Execute", recordsAffected);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="recordsAffected">object recordsAffected</param>
		/// <param name="parameters">optional object parameters</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected, object parameters)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Recordset>(this, "Execute", recordsAffected, parameters);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		/// <param name="value">optional object value</param>
		[SupportByVersion("ADODB", 2.1,2.5)]
		[BaseResult]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction, object size, object value)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter", new object[]{ name, type, direction, size, value });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter", name);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter", name, type);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter", name, type, direction);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction, object size)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.ADODBApi._Parameter>(this, "CreateParameter", name, type, direction, size);
		}

		#endregion

		#pragma warning restore
	}
}
