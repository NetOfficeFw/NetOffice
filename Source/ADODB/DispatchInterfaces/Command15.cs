using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ADODBApi
{
	///<summary>
	/// DispatchInterface Command15 
	/// SupportByVersion ADODB, 2.1,2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Command15 : _ADO
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
                    _type = typeof(Command15);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Connection ActiveConnection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveConnection", paramsArray);
				NetOffice.ADODBApi._Connection newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Connection;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ActiveConnection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public string CommandText
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandText", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandText", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public Int32 CommandTimeout
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandTimeout", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandTimeout", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public bool Prepared
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Prepared", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Prepared", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Parameters Parameters
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parameters", paramsArray);
				NetOffice.ADODBApi.Parameters newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Parameters.LateBindingApiWrapperType) as NetOffice.ADODBApi.Parameters;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi.Enums.CommandTypeEnum CommandType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CommandType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.CommandTypeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "CommandType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="recordsAffected">object RecordsAffected</param>
		/// <param name="parameters">optional object Parameters</param>
		/// <param name="options">optional Int32 Options = -1</param>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected, object parameters, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsAffected, parameters, options);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi._Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="recordsAffected">object RecordsAffected</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsAffected);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi._Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="recordsAffected">object RecordsAffected</param>
		/// <param name="parameters">optional object Parameters</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Recordset Execute(object recordsAffected, object parameters)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsAffected, parameters);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi._Recordset newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		/// <param name="value">optional object Value</param>
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction, object size, object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, direction, size, value);
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">optional string Name = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type);
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, direction);
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1, 2.5
		/// 
		/// </summary>
		/// <param name="name">optional string Name = </param>
		/// <param name="type">optional NetOffice.ADODBApi.Enums.DataTypeEnum Type = 0</param>
		/// <param name="direction">optional NetOffice.ADODBApi.Enums.ParameterDirectionEnum Direction = 1</param>
		/// <param name="size">optional Int32 Size = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.1,2.5)]
		public NetOffice.ADODBApi._Parameter CreateParameter(object name, object type, object direction, object size)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, type, direction, size);
			object returnItem = Invoker.MethodReturn(this, "CreateParameter", paramsArray);
			NetOffice.ADODBApi._Parameter newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.ADODBApi._Parameter;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}