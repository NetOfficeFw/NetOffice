using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SmartArtNode 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861178.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class SmartArtNode : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SmartArtNode
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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SmartArtNode() : base()
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
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
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
        [SupportByVersion("Office", 14, 15, 16)]
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
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.ShapeRange Shapes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ShapeRange>(this, "Shapes", typeof(NetOffice.OfficeApi.ShapeRange));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861779.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.TextFrame2 TextFrame2
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextFrame2>(this, "TextFrame2", typeof(NetOffice.OfficeApi.TextFrame2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862082.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
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
        [SupportByVersion("Office", 14, 15, 16)]
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
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.SmartArtNodes Nodes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNodes>(this, "Nodes", typeof(NetOffice.OfficeApi.SmartArtNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862047.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.SmartArtNode ParentNode
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtNode>(this, "ParentNode", typeof(NetOffice.OfficeApi.SmartArtNode));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861873.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
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
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.SmartArtNode AddNode(object position, object type)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", typeof(NetOffice.OfficeApi.SmartArtNode), position, type);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.SmartArtNode AddNode()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", typeof(NetOffice.OfficeApi.SmartArtNode));
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865366.aspx </remarks>
        /// <param name="position">optional NetOffice.OfficeApi.Enums.MsoSmartArtNodePosition Position = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public NetOffice.OfficeApi.SmartArtNode AddNode(object position)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.SmartArtNode>(this, "AddNode", typeof(NetOffice.OfficeApi.SmartArtNode), position);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863109.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862804.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Promote()
        {
            Factory.ExecuteMethod(this, "Promote");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860258.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Demote()
        {
            Factory.ExecuteMethod(this, "Demote");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864694.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Larger()
        {
            Factory.ExecuteMethod(this, "Larger");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863061.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void Smaller()
        {
            Factory.ExecuteMethod(this, "Smaller");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863035.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void ReorderUp()
        {
            Factory.ExecuteMethod(this, "ReorderUp");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860343.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public void ReorderDown()
        {
            Factory.ExecuteMethod(this, "ReorderDown");
        }

        #endregion

        #pragma warning restore
    }
}
