using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface _Index 
	/// SupportByVersion DAO, 3.6,12.0
	/// </summary>
	[SupportByVersion("DAO", 3.6,12.0)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Index : _DAO
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
                    _type = typeof(_Index);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Index(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Index(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Index(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
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

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Foreign
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Foreign");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Unique
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Unique");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Unique", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Clustered
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Clustered");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Clustered", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Required
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Required");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Required", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool IgnoreNulls
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreNulls");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreNulls", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public bool Primary
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Primary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Primary", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public Int32 DistinctCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DistinctCount");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 3.6,12.0)]
		public object Fields
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Fields");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Fields", value);
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
		public NetOffice.DAOApi.Field CreateField(object name, object type, object size)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", NetOffice.DAOApi.Field.LateBindingApiWrapperType, name, type, size);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Field CreateField()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", NetOffice.DAOApi.Field.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Field CreateField(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", NetOffice.DAOApi.Field.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Field CreateField(object name, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Field>(this, "CreateField", NetOffice.DAOApi.Field.LateBindingApiWrapperType, name, type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		/// <param name="dDL">optional object dDL</param>
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value, object dDL)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type, value, dDL);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type);
		}

		/// <summary>
		/// SupportByVersion DAO 3.6, 12.0
		/// </summary>
		/// <param name="name">optional object name</param>
		/// <param name="type">optional object type</param>
		/// <param name="value">optional object value</param>
		[CustomMethod]
		[SupportByVersion("DAO", 3.6,12.0)]
		public NetOffice.DAOApi.Property CreateProperty(object name, object type, object value)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.DAOApi.Property>(this, "CreateProperty", NetOffice.DAOApi.Property.LateBindingApiWrapperType, name, type, value);
		}

		#endregion

		#pragma warning restore
	}
}
