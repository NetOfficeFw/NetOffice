using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// XMLNamespace
    /// </summary>
    [SyntaxBypass]
    public class XMLNamespace_ : COMObject, NetOffice.WordApi.XMLNamespace_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public XMLNamespace_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Location(object allUsers)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Location", allUsers);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Location(object allUsers, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Location", allUsers, value);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_Location
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Location")]
        public virtual string Location(object allUsers)
        {
            return get_Location(allUsers);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Alias(object allUsers)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Alias", allUsers);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_Alias(object allUsers, string value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "Alias", allUsers, value);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_Alias
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Alias")]
        public virtual string Alias(object allUsers)
        {
            return get_Alias(allUsers);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.XSLTransform get_DefaultTransform(object allUsers)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransform>(this, "DefaultTransform", typeof(NetOffice.WordApi.XSLTransform), allUsers);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional XSLTransform value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_DefaultTransform(object allUsers, NetOffice.WordApi.XSLTransform value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "DefaultTransform", allUsers, value);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_DefaultTransform
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_DefaultTransform")]
        public virtual NetOffice.WordApi.XSLTransform DefaultTransform(object allUsers)
        {
            return get_DefaultTransform(allUsers);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface XMLNamespace 
    /// SupportByVersion Word, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835755.aspx </remarks>
    [SupportByVersion("Word", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class XMLNamespace : XMLNamespace_, NetOffice.WordApi.XMLNamespace
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
                    _contractType = typeof(NetOffice.WordApi.XMLNamespace);
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
                    _type = typeof(XMLNamespace);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public XMLNamespace() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192056.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196231.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193352.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195935.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string URI
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "URI");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string Location
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Location");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Location", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string Alias
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Alias");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Alias", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822199.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XSLTransforms XSLTransforms
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransforms>(this, "XSLTransforms", typeof(NetOffice.WordApi.XSLTransforms));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XSLTransform DefaultTransform
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XSLTransform>(this, "DefaultTransform", typeof(NetOffice.WordApi.XSLTransform));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "DefaultTransform", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191784.aspx </remarks>
        /// <param name="document">object document</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void AttachToDocument(object document)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AttachToDocument", document);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836059.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}
