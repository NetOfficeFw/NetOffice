using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ParagraphFormat2 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860250.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ParagraphFormat2 : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ParagraphFormat2
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
                    _type = typeof(ParagraphFormat2);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ParagraphFormat2() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862847.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862834.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoParagraphAlignment Alignment
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoParagraphAlignment>(this, "Alignment");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Alignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861479.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoBaselineAlignment BaselineAlignment
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoBaselineAlignment>(this, "BaselineAlignment");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BaselineAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863871.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.BulletFormat2 Bullet
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.BulletFormat2>(this, "Bullet", typeof(NetOffice.OfficeApi.BulletFormat2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864592.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState FarEastLineBreakLevel
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FarEastLineBreakLevel");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "FarEastLineBreakLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861781.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single FirstLineIndent
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "FirstLineIndent");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FirstLineIndent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863364.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState HangingPunctuation
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HangingPunctuation");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "HangingPunctuation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863818.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 IndentLevel
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "IndentLevel");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "IndentLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864662.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single LeftIndent
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "LeftIndent");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LeftIndent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865358.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState LineRuleAfter
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LineRuleAfter");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LineRuleAfter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863089.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState LineRuleBefore
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LineRuleBefore");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LineRuleBefore", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861150.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState LineRuleWithin
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "LineRuleWithin");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "LineRuleWithin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862758.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single RightIndent
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "RightIndent");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RightIndent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865306.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single SpaceAfter
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "SpaceAfter");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SpaceAfter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861066.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single SpaceBefore
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "SpaceBefore");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SpaceBefore", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865291.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single SpaceWithin
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "SpaceWithin");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SpaceWithin", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862162.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TabStops2 TabStops
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TabStops2>(this, "TabStops", typeof(NetOffice.OfficeApi.TabStops2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864981.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTextDirection TextDirection
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextDirection>(this, "TextDirection");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TextDirection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862056.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState WordWrap
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "WordWrap");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "WordWrap", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
