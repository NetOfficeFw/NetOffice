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
	/// DispatchInterface _Stream_Deprecated 
	/// SupportByVersion ADODB, 2.5
	///</summary>
	[SupportByVersionAttribute("ADODB", 2.5)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Stream_Deprecated : COMObject
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
                    _type = typeof(_Stream_Deprecated);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Stream_Deprecated(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream_Deprecated(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 Size
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Size", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public bool EOS
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "EOS", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public Int32 Position
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Position", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Position", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.StreamTypeEnum Type
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Type", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.StreamTypeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Type", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.LineSeparatorEnum LineSeparator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LineSeparator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.LineSeparatorEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LineSeparator", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ObjectStateEnum State
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "State", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.ObjectStateEnum)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Mode", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ADODBApi.Enums.ConnectModeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Mode", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string Charset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Charset", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Charset", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="numBytes">optional Int32 NumBytes = -1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public object Read(object numBytes)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numBytes);
			object returnItem = Invoker.MethodReturn(this, "Read", paramsArray);
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

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public object Read()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Read", paramsArray);
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

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object mode, object options, object userName, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, mode, options, userName, password);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object mode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, mode);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object mode, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, mode, options);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="source">optional object Source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Open(object source, object mode, object options, object userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(source, mode, options, userName);
			Invoker.Method(this, "Open", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void SkipLine()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SkipLine", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="buffer">object Buffer</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Write(object buffer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(buffer);
			Invoker.Method(this, "Write", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void SetEOS()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SetEOS", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream_Deprecated DestStream</param>
		/// <param name="charNumber">optional Int32 CharNumber = -1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void CopyTo(NetOffice.ADODBApi._Stream_Deprecated destStream, object charNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destStream, charNumber);
			Invoker.Method(this, "CopyTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream_Deprecated DestStream</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void CopyTo(NetOffice.ADODBApi._Stream_Deprecated destStream)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(destStream);
			Invoker.Method(this, "CopyTo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Flush()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Flush", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.SaveOptionsEnum Options = 1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void SaveToFile(string fileName, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, options);
			Invoker.Method(this, "SaveToFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void SaveToFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "SaveToFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void LoadFromFile(string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName);
			Invoker.Method(this, "LoadFromFile", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="numChars">optional Int32 NumChars = -1</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string ReadText(object numChars)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(numChars);
			object returnItem = Invoker.MethodReturn(this, "ReadText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public string ReadText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ReadText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="data">string Data</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamWriteEnum Options = 0</param>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void WriteText(string data, object options)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data, options);
			Invoker.Method(this, "WriteText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		/// <param name="data">string Data</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void WriteText(string data)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(data);
			Invoker.Method(this, "WriteText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// 
		/// </summary>
		[SupportByVersionAttribute("ADODB", 2.5)]
		public void Cancel()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cancel", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}