using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface IAllowEditRange 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
 	public class IAllowEditRange : COMObject
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
                    _type = typeof(IAllowEditRange);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IAllowEditRange(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IAllowEditRange(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IAllowEditRange(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", NetOffice.ExcelApi.Range.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Range", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.UserAccessList Users
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.UserAccessList>(this, "Users", NetOffice.ExcelApi.UserAccessList.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">string password</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 ChangePassword(string password)
		{
			return Factory.ExecuteInt32MethodGet(this, "ChangePassword", password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Unprotect(object password)
		{
			return Factory.ExecuteInt32MethodGet(this, "Unprotect", password);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 Unprotect()
		{
			return Factory.ExecuteInt32MethodGet(this, "Unprotect");
		}

		#endregion

		#pragma warning restore
	}
}
