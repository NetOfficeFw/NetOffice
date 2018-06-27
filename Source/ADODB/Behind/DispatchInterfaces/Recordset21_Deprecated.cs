using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
	/// <summary>
	/// DispatchInterface Recordset21_Deprecated 
	/// SupportByVersion ADODB, 2.5
	/// </summary>
	[SupportByVersion("ADODB", 2.5)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class Recordset21_Deprecated : Recordset20_Deprecated, NetOffice.ADODBApi.Recordset21_Deprecated
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
                    _contractType = typeof(NetOffice.ADODBApi.Recordset21_Deprecated);
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
                    _type = typeof(Recordset21_Deprecated);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Recordset21_Deprecated() : base()
		{

		}

		#endregion
		
		#region Properties

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
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object get_Collect(object index)
		{
			return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Collect", index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void set_Collect(object index, object value)
		{
			InvokerService.InvokeInternal.ExecutePropertySet(this, "Collect", index, value);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Alias for get_Collect
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("ADODB", 2.5), Redirect("get_Collect")]
		public virtual object Collect(object index)
		{
			return get_Collect(index);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// Get/Set
		/// </summary>
		[SupportByVersion("ADODB", 2.5)]
		public virtual string Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Index");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		/// <param name="seekOption">optional NetOffice.ADODBApi.Enums.SeekEnum SeekOption = 1</param>
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Seek(object keyValues, object seekOption)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", keyValues, seekOption);
		}

		/// <summary>
		/// SupportByVersion ADODB 2.5
		/// </summary>
		/// <param name="keyValues">object keyValues</param>
		[CustomMethod]
		[SupportByVersion("ADODB", 2.5)]
		public virtual void Seek(object keyValues)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Seek", keyValues);
		}

		#endregion

		#pragma warning restore
	}
}


