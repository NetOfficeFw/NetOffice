using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// XSLTransform
    /// </summary>
    [SyntaxBypass]
    public interface XSLTransform_ : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193709.aspx
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193709.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Alias")]
        string Alias(object allUsers);

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="allUsers">optional bool allUsers</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198178.aspx
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198178.aspx </remarks>
        /// <param name="allUsers">optional bool allUsers</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_Location")]
        string Location(object allUsers);

        #endregion
    }

    /// <summary>
    /// DispatchInterface XSLTransform 
    /// SupportByVersion Word, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838316.aspx </remarks>
    [SupportByVersion("Word", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
	[TypeId("E3124493-7D6A-410F-9A48-CC822C033CEC")]
    public interface XSLTransform : XSLTransform_
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198185.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838912.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193628.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193709.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string Alias { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198178.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        new string Location { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195613.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        string ID { get; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192422.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void Delete();

        #endregion
    }
}
