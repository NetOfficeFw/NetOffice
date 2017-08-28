using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Crop 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860761.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Crop : _IMsoDispObj
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
                    _type = typeof(Crop);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Crop(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Crop(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Crop(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862450.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single PictureOffsetX
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PictureOffsetX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureOffsetX", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864637.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single PictureOffsetY
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PictureOffsetY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureOffsetY", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860544.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single PictureWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PictureWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860512.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single PictureHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PictureHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PictureHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861232.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single ShapeLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ShapeLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShapeLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861517.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single ShapeTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ShapeTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShapeTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861716.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single ShapeWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ShapeWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShapeWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864643.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Single ShapeHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "ShapeHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShapeHeight", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
