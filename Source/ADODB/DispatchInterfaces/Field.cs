using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ADODBApi
{
	/// <summary>
	/// DispatchInterface Field 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Field : Field20
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
                    _type = typeof(Field);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Field(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Field(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public Int32 ActualSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ActualSize");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public Int32 Attributes
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Attributes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Attributes", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public Int32 DefinedSize
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DefinedSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefinedSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public NetOffice.ADODBApi.Enums.DataTypeEnum Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.DataTypeEnum>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public byte Precision
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "Precision");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Precision", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public byte NumericScale
		{
			get
			{
				return Factory.ExecuteBytePropertyGet(this, "NumericScale");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumericScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public object OriginalValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "OriginalValue");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public object UnderlyingValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "UnderlyingValue");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1), ProxyResult]
		public object DataFormat
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "DataFormat");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "DataFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public NetOffice.ADODBApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Properties>(this, "Properties", NetOffice.ADODBApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public Int32 Status
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Status");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="data">object data</param>
		[SupportByVersion("ADODB", 2.1)]
		public void AppendChunk(object data)
		{
			 Factory.ExecuteMethod(this, "AppendChunk", data);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="length">Int32 length</param>
		[SupportByVersion("ADODB", 2.1)]
		public object GetChunk(Int32 length)
		{
			return Factory.ExecuteVariantMethodGet(this, "GetChunk", length);
		}

		#endregion

		#pragma warning restore
	}
}
