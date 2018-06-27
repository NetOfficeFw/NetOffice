using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.DAOApi;

namespace NetOffice.DAOApi.Behind
{
	/// <summary>
	/// DispatchInterface Field2 
	/// SupportByVersion DAO, 12.0
	/// </summary>
	[SupportByVersion("DAO", 12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Field2 : _Field, NetOffice.DAOApi.Field2
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
                    _contractType = typeof(NetOffice.DAOApi.Field2);
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
                    _type = typeof(Field2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Field2() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public virtual NetOffice.DAOApi.Properties Properties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Properties>(this, "Properties", typeof(NetOffice.DAOApi.Properties));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public virtual NetOffice.DAOApi.ComplexType ComplexType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.ComplexType>(this, "ComplexType", typeof(NetOffice.DAOApi.ComplexType));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public virtual bool IsComplex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsComplex");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public virtual bool AppendOnly
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AppendOnly");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AppendOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public virtual string Expression
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Expression");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Expression", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		public virtual void LoadFromFile(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "LoadFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		public virtual void SaveToFile(string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveToFile", fileName);
		}

		#endregion

		#pragma warning restore
	}
}


