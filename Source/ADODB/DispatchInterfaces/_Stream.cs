using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface _Stream 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Stream : COMObject
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
                    _type = typeof(_Stream);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Stream(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Stream(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Stream(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Size
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public bool EOS
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EOS");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Position
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Position");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.StreamTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.StreamTypeEnum>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.LineSeparatorEnum LineSeparator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.LineSeparatorEnum>(this, "LineSeparator");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LineSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ObjectStateEnum State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ObjectStateEnum>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Enums.ConnectModeEnum Mode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.ConnectModeEnum>(this, "Mode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Mode", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public string Charset
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Charset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Charset", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numBytes">optional Int32 NumBytes = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public object Read(object numBytes)
		{
			return Factory.ExecuteVariantMethodGet(this, "Read", numBytes);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public object Read()
		{
			return Factory.ExecuteVariantMethodGet(this, "Read");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		/// <param name="password">optional string password</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object mode, object options, object userName, object password)
		{
			 Factory.ExecuteMethod(this, "Open", new object[]{ source, mode, options, userName, password });
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source)
		{
			 Factory.ExecuteMethod(this, "Open", source);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object mode)
		{
			 Factory.ExecuteMethod(this, "Open", source, mode);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object mode, object options)
		{
			 Factory.ExecuteMethod(this, "Open", source, mode, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string userName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void Open(object source, object mode, object options, object userName)
		{
			 Factory.ExecuteMethod(this, "Open", source, mode, options, userName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void SkipLine()
		{
			 Factory.ExecuteMethod(this, "SkipLine");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="buffer">object buffer</param>
		[SupportByVersion("ADODB", 2.5)]
		public void Write(object buffer)
		{
			 Factory.ExecuteMethod(this, "Write", buffer);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void SetEOS()
		{
			 Factory.ExecuteMethod(this, "SetEOS");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream destStream</param>
		/// <param name="charNumber">optional Int32 CharNumber = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public void CopyTo(NetOffice.ADODBApi._Stream destStream, object charNumber)
		{
			 Factory.ExecuteMethod(this, "CopyTo", destStream, charNumber);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream destStream</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void CopyTo(NetOffice.ADODBApi._Stream destStream)
		{
			 Factory.ExecuteMethod(this, "CopyTo", destStream);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Flush()
		{
			 Factory.ExecuteMethod(this, "Flush");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.SaveOptionsEnum Options = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		public void SaveToFile(string fileName, object options)
		{
			 Factory.ExecuteMethod(this, "SaveToFile", fileName, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void SaveToFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveToFile", fileName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("ADODB", 2.5)]
		public void LoadFromFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "LoadFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numChars">optional Int32 NumChars = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public string ReadText(object numChars)
		{
			return Factory.ExecuteStringMethodGet(this, "ReadText", numChars);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public string ReadText()
		{
			return Factory.ExecuteStringMethodGet(this, "ReadText");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamWriteEnum Options = 0</param>
		[SupportByVersion("ADODB", 2.5)]
		public void WriteText(string data, object options)
		{
			 Factory.ExecuteMethod(this, "WriteText", data, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public void WriteText(string data)
		{
			 Factory.ExecuteMethod(this, "WriteText", data);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public void Cancel()
		{
			 Factory.ExecuteMethod(this, "Cancel");
		}

		#endregion

		#pragma warning restore
	}
}
