using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// Range
    /// </summary>
    [SyntaxBypass]
    public class Range_ : COMObject, NetOffice.WordApi.Range_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public Range_() : base()
        {
        }

        #endregion

        #region Properties
        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_XML(object dataOnly)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML", dataOnly);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Alias for get_XML
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
        /// <param name="dataOnly">optional bool dataOnly</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16), Redirect("get_XML")]
        public virtual string XML(object dataOnly)
        {
            return get_XML(dataOnly);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface Range 
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845882.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Range : Range_, NetOffice.WordApi.Range
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
                    _contractType = typeof(NetOffice.WordApi.Range);
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
                    _type = typeof(Range);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public Range() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195101.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192541.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range FormattedText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "FormattedText", typeof(NetOffice.WordApi.Range));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "FormattedText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836102.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Start
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Start");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Start", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840998.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 End
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "End");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "End", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821026.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Font Font
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Font>(this, "Font", typeof(NetOffice.WordApi.Font));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Font", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837543.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Duplicate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Duplicate", typeof(NetOffice.WordApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837652.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdStoryType StoryType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStoryType>(this, "StoryType");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191956.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Tables Tables
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", typeof(NetOffice.WordApi.Tables));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836346.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Words Words
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Words>(this, "Words", typeof(NetOffice.WordApi.Words));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840991.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Sentences Sentences
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sentences>(this, "Sentences", typeof(NetOffice.WordApi.Sentences));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845462.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Characters Characters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Characters>(this, "Characters", typeof(NetOffice.WordApi.Characters));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196597.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Footnotes Footnotes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Footnotes>(this, "Footnotes", typeof(NetOffice.WordApi.Footnotes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193114.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Endnotes Endnotes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Endnotes>(this, "Endnotes", typeof(NetOffice.WordApi.Endnotes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192150.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Comments Comments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Comments>(this, "Comments", typeof(NetOffice.WordApi.Comments));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836072.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Cells Cells
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Cells>(this, "Cells", typeof(NetOffice.WordApi.Cells));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834837.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Sections Sections
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sections>(this, "Sections", typeof(NetOffice.WordApi.Sections));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837006.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Paragraphs Paragraphs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraphs>(this, "Paragraphs", typeof(NetOffice.WordApi.Paragraphs));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835448.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Borders Borders
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", typeof(NetOffice.WordApi.Borders));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Borders", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822980.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Shading Shading
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shading>(this, "Shading", typeof(NetOffice.WordApi.Shading));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839529.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.TextRetrievalMode TextRetrievalMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextRetrievalMode>(this, "TextRetrievalMode", typeof(NetOffice.WordApi.TextRetrievalMode));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "TextRetrievalMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845620.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Fields Fields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Fields>(this, "Fields", typeof(NetOffice.WordApi.Fields));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834816.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FormFields FormFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FormFields>(this, "FormFields", typeof(NetOffice.WordApi.FormFields));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837877.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Frames Frames
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frames>(this, "Frames", typeof(NetOffice.WordApi.Frames));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834843.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ParagraphFormat ParagraphFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ParagraphFormat>(this, "ParagraphFormat", typeof(NetOffice.WordApi.ParagraphFormat));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ParagraphFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195640.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ListFormat ListFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListFormat>(this, "ListFormat", typeof(NetOffice.WordApi.ListFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195181.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Bookmarks Bookmarks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bookmarks>(this, "Bookmarks", typeof(NetOffice.WordApi.Bookmarks));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196242.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839336.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822923.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844991.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Bold
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Bold");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Bold", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821583.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Italic
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Italic");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Italic", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821959.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdUnderline Underline
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdUnderline>(this, "Underline");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Underline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198151.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEmphasisMark>(this, "EmphasisMark");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EmphasisMark", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844978.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisableCharacterSpaceGrid
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisableCharacterSpaceGrid");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableCharacterSpaceGrid", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838481.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Revisions Revisions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Revisions>(this, "Revisions", typeof(NetOffice.WordApi.Revisions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836418.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Style
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Style");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Style", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845486.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StoryLength
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StoryLength");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837028.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SynonymInfo SynonymInfo
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", typeof(NetOffice.WordApi.SynonymInfo));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838128.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Hyperlinks Hyperlinks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Hyperlinks>(this, "Hyperlinks", typeof(NetOffice.WordApi.Hyperlinks));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838758.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ListParagraphs ListParagraphs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListParagraphs>(this, "ListParagraphs", typeof(NetOffice.WordApi.ListParagraphs));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837692.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Subdocuments Subdocuments
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Subdocuments>(this, "Subdocuments", typeof(NetOffice.WordApi.Subdocuments));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840317.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool GrammarChecked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "GrammarChecked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GrammarChecked", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196502.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SpellingChecked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpellingChecked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpellingChecked", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841064.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdColorIndex HighlightColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "HighlightColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "HighlightColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197474.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Columns Columns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Columns>(this, "Columns", typeof(NetOffice.WordApi.Columns));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840908.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Rows Rows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Rows>(this, "Rows", typeof(NetOffice.WordApi.Rows));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 CanEdit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CanEdit");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 CanPaste
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CanPaste");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845343.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IsEndOfRowMark
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsEndOfRowMark");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845646.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BookmarkID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BookmarkID");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191844.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 PreviousBookmarkID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PreviousBookmarkID");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195912.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Find Find
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Find>(this, "Find", typeof(NetOffice.WordApi.Find));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192629.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.PageSetup PageSetup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PageSetup>(this, "PageSetup", typeof(NetOffice.WordApi.PageSetup));
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "PageSetup", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837242.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ShapeRange ShapeRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShapeRange>(this, "ShapeRange", typeof(NetOffice.WordApi.ShapeRange));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834838.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdCharacterCase Case
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCharacterCase>(this, "Case");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Case", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_Information(NetOffice.WordApi.Enums.WdInformation type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Information", type);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_Information
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Information")]
        public virtual object Information(NetOffice.WordApi.Enums.WdInformation type)
        {
            return get_Information(type);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837707.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReadabilityStatistics>(this, "ReadabilityStatistics", typeof(NetOffice.WordApi.ReadabilityStatistics));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192406.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "GrammaticalErrors", typeof(NetOffice.WordApi.ProofreadingErrors));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195285.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ProofreadingErrors SpellingErrors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "SpellingErrors", typeof(NetOffice.WordApi.ProofreadingErrors));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195776.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdTextOrientation Orientation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTextOrientation>(this, "Orientation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Orientation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195321.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.InlineShapes InlineShapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.InlineShapes>(this, "InlineShapes", typeof(NetOffice.WordApi.InlineShapes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193730.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range NextStoryRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "NextStoryRange", typeof(NetOffice.WordApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193321.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDFarEast");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageIDFarEast", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844803.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDOther");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageIDOther", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839349.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool LanguageDetected
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "LanguageDetected");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LanguageDetected", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197181.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single FitTextWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "FitTextWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FitTextWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191976.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdHorizontalInVerticalType HorizontalInVertical
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdHorizontalInVerticalType>(this, "HorizontalInVertical");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "HorizontalInVertical", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845231.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdTwoLinesInOneType TwoLinesInOne
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTwoLinesInOneType>(this, "TwoLinesInOne");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TwoLinesInOne", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195015.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CombineCharacters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CombineCharacters");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CombineCharacters", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844920.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 NoProofing
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NoProofing");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoProofing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194640.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Tables TopLevelTables
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "TopLevelTables", typeof(NetOffice.WordApi.Tables));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192353.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Scripts Scripts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", typeof(NetOffice.OfficeApi.Scripts));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822135.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdCharacterWidth CharacterWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCharacterWidth>(this, "CharacterWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CharacterWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840112.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdKana Kana
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdKana>(this, "Kana");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Kana", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821869.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BoldBi
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BoldBi");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BoldBi", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197717.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ItalicBi
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItalicBi");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ItalicBi", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196542.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194856.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.HTMLDivisions HTMLDivisions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HTMLDivisions>(this, "HTMLDivisions", typeof(NetOffice.WordApi.HTMLDivisions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SmartTags SmartTags
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTags>(this, "SmartTags", typeof(NetOffice.WordApi.SmartTags));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820977.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowAll
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowAll");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowAll", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194311.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document Document
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "Document", typeof(NetOffice.WordApi.Document));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195199.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FootnoteOptions FootnoteOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FootnoteOptions>(this, "FootnoteOptions", typeof(NetOffice.WordApi.FootnoteOptions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195039.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.EndnoteOptions EndnoteOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EndnoteOptions>(this, "EndnoteOptions", typeof(NetOffice.WordApi.EndnoteOptions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNodes XMLNodes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLNodes", typeof(NetOffice.WordApi.XMLNodes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNode XMLParentNode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "XMLParentNode", typeof(NetOffice.WordApi.XMLNode));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840972.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Editors Editors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Editors>(this, "Editors", typeof(NetOffice.WordApi.Editors));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual string XML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XML");
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192034.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual object EnhMetaFileBits
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "EnhMetaFileBits");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822393.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.OMaths OMaths
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMaths>(this, "OMaths", typeof(NetOffice.WordApi.OMaths));
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192339.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual object CharacterStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "CharacterStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196075.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual object ParagraphStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ParagraphStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196585.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual object ListStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ListStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841045.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual object TableStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "TableStyle");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839822.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ContentControls ContentControls
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControls>(this, "ContentControls", typeof(NetOffice.WordApi.ContentControls));
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837448.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual string WordOpenXML
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WordOpenXML");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839629.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ContentControl ParentContentControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControl>(this, "ParentContentControl", typeof(NetOffice.WordApi.ContentControl));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845600.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.CoAuthLocks Locks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthLocks>(this, "Locks", typeof(NetOffice.WordApi.CoAuthLocks));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196284.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.CoAuthUpdates Updates
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthUpdates>(this, "Updates", typeof(NetOffice.WordApi.CoAuthUpdates));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823246.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Conflicts Conflicts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Conflicts>(this, "Conflicts", typeof(NetOffice.WordApi.Conflicts));
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231893.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual Int32 TextVisibleOnScreen
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextVisibleOnScreen");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820813.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Select()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823262.aspx </remarks>
        /// <param name="start">Int32 start</param>
        /// <param name="end">Int32 end</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetRange(Int32 start, Int32 end)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetRange", start, end);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
        /// <param name="direction">optional object direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Collapse(object direction)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Collapse", direction);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Collapse()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Collapse");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836272.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBefore(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBefore", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192427.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertAfter(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertAfter", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Next(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", typeof(NetOffice.WordApi.Range), unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Next()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Next(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", typeof(NetOffice.WordApi.Range), unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Previous(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", typeof(NetOffice.WordApi.Range), unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Previous()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Previous(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", typeof(NetOffice.WordApi.Range), unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartOf(object unit, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartOf", unit, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartOf()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartOf");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartOf(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartOf", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndOf(object unit, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndOf", unit, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndOf()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndOf");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndOf(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndOf", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStart(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStart", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStart()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStart");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStart(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStart", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEnd(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEnd", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEnd()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEnd");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEnd(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEnd", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveWhile(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveWhile", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveWhile(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveWhile", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStartWhile(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStartWhile", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStartWhile(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStartWhile", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEndWhile(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEndWhile", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEndWhile(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEndWhile", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUntil(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUntil", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUntil(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUntil", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStartUntil(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStartUntil", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStartUntil(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStartUntil", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
        /// <param name="cset">object cset</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEndUntil(object cset, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEndUntil", cset, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
        /// <param name="cset">object cset</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEndUntil(object cset)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEndUntil", cset);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195686.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837718.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845105.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBreak(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBreak", type);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBreak()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBreak");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        /// <param name="link">optional object link</param>
        /// <param name="attachment">optional object attachment</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFile(string fileName, object range, object confirmConversions, object link, object attachment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFile", new object[] { fileName, range, confirmConversions, link, attachment });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFile(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFile", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFile(string fileName, object range)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFile", fileName, range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFile(string fileName, object range, object confirmConversions)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFile", fileName, range, confirmConversions);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="range">optional object range</param>
        /// <param name="confirmConversions">optional object confirmConversions</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFile(string fileName, object range, object confirmConversions, object link)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFile", fileName, range, confirmConversions, link);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197125.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool InStory(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "InStory", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822960.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool InRange(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "InRange", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837449.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void WholeStory()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "WholeStory");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Expand(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Expand", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Expand()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Expand");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196197.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertParagraph()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertParagraph");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822546.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertParagraphAfter()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertParagraphAfter");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), separator);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), separator, numRows);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), separator, numRows, numColumns);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), separator, numRows, numColumns, initialColumnWidth);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTimeOld(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat, insertAsField, insertAsFullWidth);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTimeOld()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTimeOld");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTimeOld(object dateTimeFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTimeOld(object dateTimeFormat, object insertAsField)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat, insertAsField);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        /// <param name="bias">optional object bias</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSymbol", characterNumber, font, unicode, bias);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertSymbol(Int32 characterNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSymbol", characterNumber);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertSymbol(Int32 characterNumber, object font)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSymbol", characterNumber, font);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
        /// <param name="characterNumber">Int32 characterNumber</param>
        /// <param name="font">optional object font</param>
        /// <param name="unicode">optional object unicode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertSymbol(Int32 characterNumber, object font, object unicode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertSymbol", characterNumber, font, unicode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference", new object[] { referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition });
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        /// <param name="separateNumbers">optional object separateNumbers</param>
        /// <param name="separatorString">optional object separatorString</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers, object separatorString)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference", new object[] { referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers, separatorString });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference", referenceType, referenceKind, referenceItem);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference", referenceType, referenceKind, referenceItem, insertAsHyperlink);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        /// <param name="separateNumbers">optional object separateNumbers</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference", new object[] { referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCaption(object label, object title, object titleAutoText, object position)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaption", label, title, titleAutoText, position);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        /// <param name="excludeLabel">optional object excludeLabel</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCaption(object label, object title, object titleAutoText, object position, object excludeLabel)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaption", new object[] { label, title, titleAutoText, position, excludeLabel });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCaption(object label)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaption", label);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCaption(object label, object title)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaption", label, title);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCaption(object label, object title, object titleAutoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaption", label, title, titleAutoText);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836633.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CopyAsPicture()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyAsPicture");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="languageID">optional object languageID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object languageID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, languageID });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortOld", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193013.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortAscending()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortAscending");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844858.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortDescending()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortDescending");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838323.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IsEqual(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsEqual", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821015.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Calculate()
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "Calculate");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        /// <param name="name">optional object name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which, count, name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo(object what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo(object what, object which)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
        /// <param name="what">optional object what</param>
        /// <param name="which">optional object which</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo(object what, object which, object count)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range), what, which, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844826.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToNext", typeof(NetOffice.WordApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836673.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToPrevious", typeof(NetOffice.WordApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        /// <param name="iconLabel">optional object iconLabel</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName, object iconLabel)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", new object[] { iconIndex, link, placement, displayAsIcon, dataType, iconFileName, iconLabel });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", iconIndex);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", iconIndex, link);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link, object placement)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", iconIndex, link, placement);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", iconIndex, link, placement, displayAsIcon);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", new object[] { iconIndex, link, placement, displayAsIcon, dataType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
        /// <param name="iconIndex">optional object iconIndex</param>
        /// <param name="link">optional object link</param>
        /// <param name="placement">optional object placement</param>
        /// <param name="displayAsIcon">optional object displayAsIcon</param>
        /// <param name="dataType">optional object dataType</param>
        /// <param name="iconFileName">optional object iconFileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial", new object[] { iconIndex, link, placement, displayAsIcon, dataType, iconFileName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835691.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void LookupNameProperties()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LookupNameProperties");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196924.aspx </remarks>
        /// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192827.aspx </remarks>
        /// <param name="direction">Int32 direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Relocate(Int32 direction)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Relocate", direction);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839497.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSynonyms()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSynonyms");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        /// <param name="format">optional object format</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SubscribeTo(string edition, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SubscribeTo", edition, format);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">string edition</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SubscribeTo(string edition)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SubscribeTo", edition);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        /// <param name="containsText">optional object containsText</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreatePublisher(object edition, object containsPICT, object containsRTF, object containsText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePublisher", edition, containsPICT, containsRTF, containsText);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreatePublisher()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePublisher");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreatePublisher(object edition)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePublisher", edition);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreatePublisher(object edition, object containsPICT)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePublisher", edition, containsPICT);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="edition">optional object edition</param>
        /// <param name="containsPICT">optional object containsPICT</param>
        /// <param name="containsRTF">optional object containsRTF</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreatePublisher(object edition, object containsPICT, object containsRTF)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreatePublisher", edition, containsPICT, containsRTF);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838952.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertAutoText()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertAutoText");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="includeFields">optional object includeFields</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to, object includeFields)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to, includeFields });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", format);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", format, style);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", format, style, linkToSource);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", format, style, linkToSource, connection);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
        /// <param name="format">optional object format</param>
        /// <param name="style">optional object style</param>
        /// <param name="linkToSource">optional object linkToSource</param>
        /// <param name="connection">optional object connection</param>
        /// <param name="sQLStatement">optional object sQLStatement</param>
        /// <param name="sQLStatement1">optional object sQLStatement1</param>
        /// <param name="passwordDocument">optional object passwordDocument</param>
        /// <param name="passwordTemplate">optional object passwordTemplate</param>
        /// <param name="writePasswordDocument">optional object writePasswordDocument</param>
        /// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
        /// <param name="dataSource">optional object dataSource</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDatabase", new object[] { format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845283.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AutoFormat()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193931.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckGrammar()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckGrammar");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="alwaysSuggest">optional object alwaysSuggest</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CheckSpelling", new object[] { customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        /// <param name="customDictionary10">optional object customDictionary10</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), customDictionary, ignoreUppercase, mainDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), customDictionary, ignoreUppercase, mainDictionary, suggestionMode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        /// <param name="customDictionary9">optional object customDictionary9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821256.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertParagraphBefore()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertParagraphBefore");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195326.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void NextSubdocument()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NextSubdocument");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195945.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PreviousSubdocument()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PreviousSubdocument");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        /// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering, object customDictionary)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja", new object[] { conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering, customDictionary });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja(object conversionsMode)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja(object conversionsMode, object fastConversion)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion, checkHangulEnding);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
        /// <param name="conversionsMode">optional object conversionsMode</param>
        /// <param name="fastConversion">optional object fastConversion</param>
        /// <param name="checkHangulEnding">optional object checkHangulEnding</param>
        /// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822962.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAsNestedTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAsNestedTable");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        /// <param name="symbol">optional object symbol</param>
        /// <param name="enclosedText">optional object enclosedText</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ModifyEnclosure(object style, object symbol, object enclosedText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ModifyEnclosure", style, symbol, enclosedText);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ModifyEnclosure(object style)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ModifyEnclosure", style);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
        /// <param name="style">object style</param>
        /// <param name="symbol">optional object symbol</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ModifyEnclosure(object style, object symbol)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ModifyEnclosure", style, symbol);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        /// <param name="fontSize">optional Int32 FontSize = 0</param>
        /// <param name="fontName">optional string FontName = </param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PhoneticGuide(string text, object alignment, object raise, object fontSize, object fontName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PhoneticGuide", new object[] { text, alignment, raise, fontSize, fontName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PhoneticGuide(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PhoneticGuide", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PhoneticGuide(string text, object alignment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PhoneticGuide", text, alignment);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PhoneticGuide(string text, object alignment, object raise)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PhoneticGuide", text, alignment, raise);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
        /// <param name="text">string text</param>
        /// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
        /// <param name="raise">optional Int32 Raise = 0</param>
        /// <param name="fontSize">optional Int32 FontSize = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PhoneticGuide(string text, object alignment, object raise, object fontSize)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PhoneticGuide", text, alignment, raise, fontSize);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        /// <param name="dateLanguage">optional object dateLanguage</param>
        /// <param name="calendarType">optional object calendarType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage, object calendarType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime", new object[] { dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage, calendarType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime(object dateTimeFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime", dateTimeFormat);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime(object dateTimeFormat, object insertAsField)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField, insertAsFullWidth);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
        /// <param name="dateTimeFormat">optional object dateTimeFormat</param>
        /// <param name="insertAsField">optional object insertAsField</param>
        /// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
        /// <param name="dateLanguage">optional object dateLanguage</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        /// <param name="sortOrder2">optional object sortOrder2</param>
        /// <param name="fieldNumber3">optional object fieldNumber3</param>
        /// <param name="sortFieldType3">optional object sortFieldType3</param>
        /// <param name="sortOrder3">optional object sortOrder3</param>
        /// <param name="sortColumn">optional object sortColumn</param>
        /// <param name="separator">optional object separator</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195289.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DetectLanguage()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DetectLanguage");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        /// <param name="autoFitBehavior">optional object autoFitBehavior</param>
        /// <param name="defaultTableBehavior">optional object defaultTableBehavior</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior, object defaultTableBehavior)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior, defaultTableBehavior });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), separator);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), separator, numRows);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), separator, numRows, numColumns);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), separator, numRows, numColumns, initialColumnWidth);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
        /// <param name="separator">optional object separator</param>
        /// <param name="numRows">optional object numRows</param>
        /// <param name="numColumns">optional object numColumns</param>
        /// <param name="initialColumnWidth">optional object initialColumnWidth</param>
        /// <param name="format">optional object format</param>
        /// <param name="applyBorders">optional object applyBorders</param>
        /// <param name="applyShading">optional object applyShading</param>
        /// <param name="applyFont">optional object applyFont</param>
        /// <param name="applyColor">optional object applyColor</param>
        /// <param name="applyHeadingRows">optional object applyHeadingRows</param>
        /// <param name="applyLastRow">optional object applyLastRow</param>
        /// <param name="applyFirstColumn">optional object applyFirstColumn</param>
        /// <param name="applyLastColumn">optional object applyLastColumn</param>
        /// <param name="autoFit">optional object autoFit</param>
        /// <param name="autoFitBehavior">optional object autoFitBehavior</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table), new object[] { separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        /// <param name="commonTerms">optional bool CommonTerms = false</param>
        /// <param name="useVariants">optional bool UseVariants = false</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TCSCConverter(object wdTCSCConverterDirection, object commonTerms, object useVariants)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection, commonTerms, useVariants);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TCSCConverter()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TCSCConverter");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TCSCConverter(object wdTCSCConverterDirection)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
        /// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
        /// <param name="commonTerms">optional bool CommonTerms = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TCSCConverter(object wdTCSCConverterDirection, object commonTerms)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection, commonTerms);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193749.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAndFormat", type);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193063.aspx </remarks>
        /// <param name="linkedToExcel">bool linkedToExcel</param>
        /// <param name="wordFormatting">bool wordFormatting</param>
        /// <param name="rTF">bool rTF</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteExcelTable", linkedToExcel, wordFormatting, rTF);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839173.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAppendTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAppendTable");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        /// <param name="includePosition">optional object includePosition</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference_2002", new object[] { referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition });
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference_2002", referenceType, referenceKind, referenceItem);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="referenceType">object referenceType</param>
        /// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
        /// <param name="referenceItem">object referenceItem</param>
        /// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCrossReference_2002", referenceType, referenceKind, referenceItem, insertAsHyperlink);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        /// <param name="position">optional object position</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCaptionXP(object label, object title, object titleAutoText, object position)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaptionXP", label, title, titleAutoText, position);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCaptionXP(object label)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaptionXP", label);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCaptionXP(object label, object title)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaptionXP", label, title);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="label">object label</param>
        /// <param name="title">optional object title</param>
        /// <param name="titleAutoText">optional object titleAutoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertCaptionXP(object label, object title, object titleAutoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCaptionXP", label, title, titleAutoText);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
        /// <param name="editorID">optional object editorID</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToEditableRange(object editorID)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", typeof(NetOffice.WordApi.Range), editorID);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToEditableRange()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
        /// <param name="xML">string xML</param>
        /// <param name="transform">optional object transform</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertXML(string xML, object transform)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertXML", xML, transform);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
        /// <param name="xML">string xML</param>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void InsertXML(string xML)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertXML", xML);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822335.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="format">NetOffice.WordApi.Enums.WdSaveFormat format</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportFragment(string fileName, NetOffice.WordApi.Enums.WdSaveFormat format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportFragment", fileName, format);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821878.aspx </remarks>
        /// <param name="level">Int16 level</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void SetListLevel(Int16 level)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetListLevel", level);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
        /// <param name="alignment">Int32 alignment</param>
        /// <param name="relativeTo">optional Int32 RelativeTo = 0</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void InsertAlignmentTab(Int32 alignment, object relativeTo)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertAlignmentTab", alignment, relativeTo);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
        /// <param name="alignment">Int32 alignment</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void InsertAlignmentTab(Int32 alignment)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertAlignmentTab", alignment);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="matchDestination">optional bool MatchDestination = false</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ImportFragment(string fileName, object matchDestination)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ImportFragment", fileName, matchDestination);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ImportFragment(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ImportFragment", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        /// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport, optimizeFor);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
        /// <param name="outputFileName">string outputFileName</param>
        /// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
        /// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
        /// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
        /// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
        /// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
        /// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
        /// <param name="keepIRM">optional bool KeepIRM = true</param>
        /// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
        /// <param name="docStructureTags">optional bool DocStructureTags = true</param>
        /// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
        /// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportAsFixedFormat", new object[] { outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1 });
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", new object[] { sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", sortFieldType);
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder, caseSensitive);
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder, caseSensitive, bidiSort);
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", new object[] { sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe });
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", new object[] { sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", new object[] { sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="caseSensitive">optional object caseSensitive</param>
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings", new object[] { sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
        }

        #endregion

        #pragma warning restore
    }
}

