using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Field 
	/// SupportByVersion ADODB, 2.1,2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.1,2.5)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Field : Field20, NetOffice.ADODBApi.Field
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
                    _contractType = typeof(NetOffice.ADODBApi.Field);
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
                    _type = typeof(Field);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Field() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual Int32 ActualSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ActualSize");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual Int32 Attributes
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Attributes");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Attributes", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual Int32 DefinedSize
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DefinedSize");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefinedSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual NetOffice.ADODBApi.Enums.DataTypeEnum Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ADODBApi.Enums.DataTypeEnum>(this, "Type");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual object Value
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual byte Precision
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "Precision");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Precision", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual byte NumericScale
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "NumericScale");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NumericScale", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual object OriginalValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OriginalValue");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.1)]
		public virtual object UnderlyingValue
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "UnderlyingValue");
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("ADODB", 2.1), ProxyResult]
		public virtual object DataFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "DataFormat");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DataFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual NetOffice.ADODBApi.Properties Properties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Properties>(this, "Properties", typeof(NetOffice.ADODBApi.Properties));
			}
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual Int32 Status
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Status");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="data">object data</param>
		[SupportByVersion("ADODB", 2.1)]
		public virtual void AppendChunk(object data)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AppendChunk", data);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.1
		/// </summary>
		/// <param name="length">Int32 length</param>
		[SupportByVersion("ADODB", 2.1)]
		public virtual object GetChunk(Int32 length)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "GetChunk", length);
		}

		#endregion

		#pragma warning restore
	}
}


