using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface _Stream_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class _Stream_Deprecated : COMObject, NetOffice.ADODBApi._Stream_Deprecated
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
                    _contractType = typeof(NetOffice.ADODBApi._Stream_Deprecated);
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
                    _type = typeof(_Stream_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Stream_Deprecated() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual bool EOS
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "EOS");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 Position
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Position");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.StreamTypeEnum Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.StreamTypeEnum>(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Enums.LineSeparatorEnum LineSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.LineSeparatorEnum>(this, "LineSeparator");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LineSeparator", value);
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string Charset
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Charset");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Charset", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numBytes">optional Int32 NumBytes = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual object Read(object numBytes)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Read", numBytes);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual object Read()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Read");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		/// <param name="password">optional string Password = </param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object mode, object options, object userName, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", new object[]{ source, mode, options, userName, password });
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
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object mode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, mode);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object mode, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, mode, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="source">optional object source</param>
		/// <param name="mode">optional NetOffice.ADODBApi.Enums.ConnectModeEnum Mode = 0</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamOpenOptionsEnum Options = -1</param>
		/// <param name="userName">optional string UserName = </param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Open(object source, object mode, object options, object userName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Open", source, mode, options, userName);
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
		public virtual void SkipLine()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SkipLine");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="buffer">object buffer</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Write(object buffer)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Write", buffer);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void SetEOS()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetEOS");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream_Deprecated destStream</param>
		/// <param name="charNumber">optional Int32 CharNumber = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void CopyTo(NetOffice.ADODBApi._Stream_Deprecated destStream, object charNumber)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyTo", destStream, charNumber);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="destStream">NetOffice.ADODBApi._Stream_Deprecated destStream</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void CopyTo(NetOffice.ADODBApi._Stream_Deprecated destStream)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CopyTo", destStream);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Flush()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Flush");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.SaveOptionsEnum Options = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void SaveToFile(string fileName, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveToFile", fileName, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void SaveToFile(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveToFile", fileName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void LoadFromFile(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LoadFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="numChars">optional Int32 NumChars = -1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string ReadText(object numChars)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ReadText", numChars);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual string ReadText()
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ReadText");
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		/// <param name="options">optional NetOffice.ADODBApi.Enums.StreamWriteEnum Options = 0</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void WriteText(string data, object options)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WriteText", data, options);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="data">string data</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void WriteText(string data)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WriteText", data);
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

