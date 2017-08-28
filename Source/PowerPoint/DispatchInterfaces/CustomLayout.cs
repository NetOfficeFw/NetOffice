using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface CustomLayout 
	/// SupportByVersion PowerPoint, 12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745128.aspx </remarks>
	[SupportByVersion("PowerPoint", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CustomLayout : COMObject
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
                    _type = typeof(CustomLayout);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public CustomLayout(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public CustomLayout(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public CustomLayout(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745630.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", NetOffice.PowerPointApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744548.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746646.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shapes Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shapes>(this, "Shapes", NetOffice.PowerPointApi.Shapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746590.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.HeadersFooters HeadersFooters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.HeadersFooters>(this, "HeadersFooters", NetOffice.PowerPointApi.HeadersFooters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745044.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ShapeRange>(this, "Background", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744596.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
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
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745358.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single Height
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Height");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746291.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Single Width
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Width");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745434.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Hyperlinks>(this, "Hyperlinks", NetOffice.PowerPointApi.Hyperlinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745689.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Design Design
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Design>(this, "Design", NetOffice.PowerPointApi.Design.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745079.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.TimeLine TimeLine
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.TimeLine>(this, "TimeLine", NetOffice.PowerPointApi.TimeLine.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746695.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.SlideShowTransition SlideShowTransition
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.SlideShowTransition>(this, "SlideShowTransition", NetOffice.PowerPointApi.SlideShowTransition.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746832.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public string MatchingName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MatchingName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchingName", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745529.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Preserved
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Preserved");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Preserved", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746123.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public Int32 Index
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Index");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743926.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DisplayMasterShapes
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DisplayMasterShapes");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DisplayMasterShapes", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745678.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState FollowMasterBackground
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FollowMasterBackground");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FollowMasterBackground", value);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746208.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.OfficeApi.ThemeColorScheme ThemeColorScheme
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ThemeColorScheme>(this, "ThemeColorScheme", NetOffice.OfficeApi.ThemeColorScheme.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745845.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.CustomerData CustomerData
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.CustomerData>(this, "CustomerData", NetOffice.PowerPointApi.CustomerData.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227860.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Guides Guides
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Guides>(this, "Guides", NetOffice.PowerPointApi.Guides.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744500.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743915.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746514.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745163.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746200.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.CustomLayout Duplicate()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.CustomLayout>(this, "Duplicate", NetOffice.PowerPointApi.CustomLayout.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744015.aspx </remarks>
		/// <param name="toPos">Int32 toPos</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public void MoveTo(Int32 toPos)
		{
			 Factory.ExecuteMethod(this, "MoveTo", toPos);
		}

		#endregion

		#pragma warning restore
	}
}
