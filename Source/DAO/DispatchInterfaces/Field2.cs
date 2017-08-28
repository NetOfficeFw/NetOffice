using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.DAOApi
{
	/// <summary>
	/// DispatchInterface Field2 
	/// SupportByVersion DAO, 12.0
	/// </summary>
	[SupportByVersion("DAO", 12.0)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Field2 : _Field
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
                    _type = typeof(Field2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Field2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Field2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Field2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public NetOffice.DAOApi.Properties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.Properties>(this, "Properties", NetOffice.DAOApi.Properties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public NetOffice.DAOApi.ComplexType ComplexType
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.DAOApi.ComplexType>(this, "ComplexType", NetOffice.DAOApi.ComplexType.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public bool IsComplex
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsComplex");
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public bool AppendOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AppendOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AppendOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// Get/Set
		/// </summary>
		[SupportByVersion("DAO", 12.0)]
		public string Expression
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Expression");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Expression", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		public void LoadFromFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "LoadFromFile", fileName);
		}

		/// <summary>
		/// SupportByVersion DAO 12.0
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("DAO", 12.0)]
		public void SaveToFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "SaveToFile", fileName);
		}

		#endregion

		#pragma warning restore
	}
}
