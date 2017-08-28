using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface SmartArtNode 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861178.aspx </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SmartArtNode : _IMsoDispObj
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
                    _type = typeof(SmartArtNode);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SmartArtNode(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SmartArtNode(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SmartArtNode(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863308.aspx </remarks>
		[SupportByVersion("Office", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860568.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType OrgChartLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoOrgChartLayoutType>(this, "OrgChartLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OrgChartLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864604.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.ShapeRange Shapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ShapeRange>(this, "Shapes", NetOffice.OfficeApi.ShapeRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861779.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.TextFrame2 TextFrame2
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextFrame2>(this, "TextFrame2", NetOffice.OfficeApi.TextFrame2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862082.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public Int32 Level
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Level");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860275.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Hidden
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Hidden");
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865275.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNodes Nodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNodes>(this, "Nodes", NetOffice.OfficeApi.SmartArtNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862047.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNode ParentNode
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNode>(this, "ParentNode", NetOffice.OfficeApi.SmartArtNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861873.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoSmartArtNodeType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSmartArtNodeType>(this, "Type");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
		/// <param name="position">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodePosition Position = 1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodeType Type = 1</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNode AddNode(object position, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", NetOffice.OfficeApi.SmartArtNode.LateBindingApiWrapperType, position, type);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNode AddNode()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", NetOffice.OfficeApi.SmartArtNode.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
		/// <param name="position">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodePosition Position = 1</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.SmartArtNode AddNode(object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", NetOffice.OfficeApi.SmartArtNode.LateBindingApiWrapperType, position);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863109.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862804.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Promote()
		{
			 Factory.ExecuteMethod(this, "Promote");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860258.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Demote()
		{
			 Factory.ExecuteMethod(this, "Demote");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864694.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Larger()
		{
			 Factory.ExecuteMethod(this, "Larger");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863061.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void Smaller()
		{
			 Factory.ExecuteMethod(this, "Smaller");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863035.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void ReorderUp()
		{
			 Factory.ExecuteMethod(this, "ReorderUp");
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860343.aspx </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public void ReorderDown()
		{
			 Factory.ExecuteMethod(this, "ReorderDown");
		}

		#endregion

		#pragma warning restore
	}
}
