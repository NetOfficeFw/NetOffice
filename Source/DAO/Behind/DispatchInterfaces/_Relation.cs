using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.DAOApi;

namespace NetOffice.DAOApi.Behind
{
	/// <summary>
	/// DispatchInterface _Relation 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Relation : _DAO, NetOffice.DAOApi._Relation
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
                    _contractType = typeof(NetOffice.DAOApi._Relation);
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
                    _type = typeof(_Relation);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Relation() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
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

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string Table
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Table");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Table", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual string ForeignTable
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ForeignTable");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ForeignTable", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
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
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Fields Fields
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Fields>(this, "Fields", typeof(NetOffice.DAOApi.Fields));
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual bool PartialReplica
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PartialReplica");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PartialReplica", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="size">optional object size</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Field CreateField(object name, object type, object size)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", typeof(NetOffice.DAOApi.Field), name, type, size);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Field CreateField()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", typeof(NetOffice.DAOApi.Field));
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Field CreateField(object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", typeof(NetOffice.DAOApi.Field), name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public virtual NetOffice.DAOApi.Field CreateField(object name, object type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", typeof(NetOffice.DAOApi.Field), name, type);
		}

		#endregion

		#pragma warning restore
	}
}


