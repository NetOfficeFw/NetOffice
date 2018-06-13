using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// XMLNamespace
    /// </summary>
    [SyntaxBypass]
    public interface XMLNamespace_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Location(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Location(object allUsers, string value);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_Location
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Location")]
        string Location(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string get_Alias(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional string value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_Alias(object allUsers, string value);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_Alias
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Alias")]
        string Alias(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.XSLTransform get_DefaultTransform(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// <param name="value">optional XSLTransform value</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void set_DefaultTransform(object allUsers, NetOffice.WordApi.XSLTransform value);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_DefaultTransform
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_DefaultTransform")]
        NetOffice.WordApi.XSLTransform DefaultTransform(object allUsers);

        #endregion
    }
 
    /// <summary>
    /// DispatchInterface XMLNamespace 
    /// SupportByVersion Word, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835755.aspx </remarks>
    [SupportByVersion("Word", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("B140A023-4850-4DA6-BC5F-CC459C4507BC")]
    public interface XMLNamespace : XMLNamespace_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192056.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196231.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193352.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195935.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        string URI { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192802.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string Location { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194042.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string Alias { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822199.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XSLTransforms XSLTransforms { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835718.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new NetOffice.WordApi.XSLTransform DefaultTransform { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191784.aspx </remarks>
        /// <param name="document">object document</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void AttachToDocument(object document);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836059.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
