using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// Selection
    /// </summary>
    [SyntaxBypass]
    public class Selection_ : COMObject, NetOffice.WordApi.Selection_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public Selection_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="dataOnly">optional bool dataOnly</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
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
    /// DispatchInterface Selection 
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821411.aspx </remarks>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Selection : Selection_, NetOffice.WordApi.Selection
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
                    _contractType = typeof(NetOffice.WordApi.Selection);
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
                    _type = typeof(Selection);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public Selection() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192754.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836670.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839485.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834869.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837859.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821048.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdSelectionType Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSelectionType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191739.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838978.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845908.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837460.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193720.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196946.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197009.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841006.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823219.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195296.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836277.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840393.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193012.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192021.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845839.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838307.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193858.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197430.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193356.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836357.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838983.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196398.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191830.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838134.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194663.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821842.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836744.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.HeaderFooter HeaderFooter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HeaderFooter>(this, "HeaderFooter", typeof(NetOffice.WordApi.HeaderFooter));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845161.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840519.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193388.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197434.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845594.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Range
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_Information")]
        public virtual object Information(NetOffice.WordApi.Enums.WdInformation type)
        {
            return get_Information(type);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837479.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdSelectionFlags Flags
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSelectionFlags>(this, "Flags");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Flags", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835497.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Active
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Active");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820824.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool StartIsActive
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "StartIsActive");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartIsActive", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822970.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IPAtEndOfLine
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IPAtEndOfLine");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821400.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ExtendMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ExtendMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ExtendMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839310.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ColumnSelectMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ColumnSelectMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnSelectMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821992.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193084.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192167.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196980.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839166.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844964.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document Document
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "Document", typeof(NetOffice.WordApi.Document));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836759.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196937.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821380.aspx </remarks>
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
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192601.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821699.aspx </remarks>
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198226.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191940.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ShapeRange ChildShapeRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShapeRange>(this, "ChildShapeRange", typeof(NetOffice.WordApi.ShapeRange));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191804.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool HasChildShapeRange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasChildShapeRange");
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845098.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192368.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837314.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840039.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838161.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820971.aspx </remarks>
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
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ContentControl ParentContentControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControl>(this, "ParentContentControl", typeof(NetOffice.WordApi.ContentControl));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845714.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Select()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192352.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
        /// <param name="direction">optional object direction</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Collapse(object direction)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Collapse", direction);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Collapse()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Collapse");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845077.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBefore(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBefore", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192184.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertAfter(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertAfter", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Next()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range Previous()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 StartOf()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "StartOf");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndOf()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndOf");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Move()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Move");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveStart()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveStart");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveEnd()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveEnd");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192037.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196538.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840284.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBreak(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBreak", type);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertBreak()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertBreak");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192633.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool InStory(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "InStory", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193660.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool InRange(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "InRange", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Expand(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Expand", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Expand()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Expand");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837485.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertParagraph()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertParagraph");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836408.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840576.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821863.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortAscending()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortAscending");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845052.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SortDescending()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortDescending");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196258.aspx </remarks>
        /// <param name="range">NetOffice.WordApi.Range range</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool IsEqual(NetOffice.WordApi.Range range)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsEqual", range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835748.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Calculate()
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "Calculate");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoTo()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836451.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToNext", typeof(NetOffice.WordApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839107.aspx </remarks>
        /// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToPrevious", typeof(NetOffice.WordApi.Range), what);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteSpecial()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteSpecial");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834516.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Field PreviousField()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "PreviousField", typeof(NetOffice.WordApi.Field));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194299.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Field NextField()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "NextField", typeof(NetOffice.WordApi.Field));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840515.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertParagraphBefore()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertParagraphBefore");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
        /// <param name="shiftCells">optional object shiftCells</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCells(object shiftCells)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCells", shiftCells);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertCells()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertCells");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
        /// <param name="character">optional object character</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Extend(object character)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Extend", character);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Extend()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Extend");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840081.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Shrink()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Shrink");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveLeft(object unit, object count, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveLeft", unit, count, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveLeft()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveLeft");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveLeft(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveLeft", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveLeft(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveLeft", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveRight(object unit, object count, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveRight", unit, count, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveRight()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveRight");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveRight(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveRight", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveRight(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveRight", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUp(object unit, object count, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUp", unit, count, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUp()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUp");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUp(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUp", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveUp(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveUp", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveDown(object unit, object count, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveDown", unit, count, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveDown()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveDown");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveDown(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveDown", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="count">optional object count</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 MoveDown(object unit, object count)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "MoveDown", unit, count);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 HomeKey(object unit, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HomeKey", unit, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 HomeKey()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HomeKey");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 HomeKey(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HomeKey", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        /// <param name="extend">optional object extend</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndKey(object unit, object extend)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndKey", unit, extend);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndKey()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndKey");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
        /// <param name="unit">optional object unit</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 EndKey(object unit)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "EndKey", unit);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835736.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void EscapeKey()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "EscapeKey");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840867.aspx </remarks>
        /// <param name="text">string text</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TypeText(string text)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TypeText", text);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840230.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CopyFormat()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyFormat");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196637.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteFormat()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteFormat");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839799.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TypeParagraph()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TypeParagraph");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194909.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void TypeBackspace()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "TypeBackspace");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839790.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void NextSubdocument()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NextSubdocument");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845750.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PreviousSubdocument()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PreviousSubdocument");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836022.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectColumn()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectColumn");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197469.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentFont()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentFont");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822643.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentAlignment()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentAlignment");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191872.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentSpacing()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentSpacing");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193883.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentIndent()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentIndent");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193718.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentTabs()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentTabs");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840690.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCurrentColor()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCurrentColor");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839540.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void CreateTextbox()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CreateTextbox");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840046.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void WholeStory()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "WholeStory");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845469.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectRow()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectRow");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196707.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SplitTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SplitTable");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRows(object numRows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRows", numRows);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRows()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRows");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838759.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertColumns()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertColumns");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        /// <param name="formula">optional object formula</param>
        /// <param name="numberFormat">optional object numberFormat</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFormula(object formula, object numberFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFormula", formula, numberFormat);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFormula()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFormula");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
        /// <param name="formula">optional object formula</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertFormula(object formula)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertFormula", formula);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
        /// <param name="wrap">optional object wrap</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Revision NextRevision(object wrap)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "NextRevision", typeof(NetOffice.WordApi.Revision), wrap);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Revision NextRevision()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "NextRevision", typeof(NetOffice.WordApi.Revision));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
        /// <param name="wrap">optional object wrap</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Revision PreviousRevision(object wrap)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "PreviousRevision", typeof(NetOffice.WordApi.Revision), wrap);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Revision PreviousRevision()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "PreviousRevision", typeof(NetOffice.WordApi.Revision));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194535.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAsNestedTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAsNestedTable");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839331.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="styleName">string styleName</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.AutoTextEntry CreateAutoTextEntry(string name, string styleName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.AutoTextEntry>(this, "CreateAutoTextEntry", typeof(NetOffice.WordApi.AutoTextEntry), name, styleName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838494.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DetectLanguage()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DetectLanguage");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195143.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SelectCell()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SelectCell");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRowsBelow(object numRows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRowsBelow", numRows);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRowsBelow()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRowsBelow");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844950.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertColumnsRight()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertColumnsRight");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
        /// <param name="numRows">optional object numRows</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRowsAbove(object numRows)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRowsAbove", numRows);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertRowsAbove()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertRowsAbove");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821034.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RtlRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RtlRun");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839502.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void LtrRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LtrRun");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845275.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void BoldRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "BoldRun");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845442.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ItalicRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ItalicRun");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836904.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RtlPara()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RtlPara");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834853.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void LtrPara()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LtrPara");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void InsertDateTime()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertDateTime");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        /// <param name="subFieldNumber2">optional object subFieldNumber2</param>
        /// <param name="subFieldNumber3">optional object subFieldNumber3</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2, object subFieldNumber3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2, subFieldNumber3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Sort()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
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
        /// <param name="subFieldNumber">optional object subFieldNumber</param>
        /// <param name="subFieldNumber2">optional object subFieldNumber2</param>
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Table ConvertToTable()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", typeof(NetOffice.WordApi.Table));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
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
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        /// <param name="languageID">optional object languageID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", excludeHeader);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber, sortFieldType);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber, sortFieldType, sortOrder);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="excludeHeader">optional object excludeHeader</param>
        /// <param name="fieldNumber">optional object fieldNumber</param>
        /// <param name="sortFieldType">optional object sortFieldType</param>
        /// <param name="sortOrder">optional object sortOrder</param>
        /// <param name="fieldNumber2">optional object fieldNumber2</param>
        /// <param name="sortFieldType2">optional object sortFieldType2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
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
        /// <param name="bidiSort">optional object bidiSort</param>
        /// <param name="ignoreThe">optional object ignoreThe</param>
        /// <param name="ignoreKashida">optional object ignoreKashida</param>
        /// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
        /// <param name="ignoreHe">optional object ignoreHe</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Sort2000", new object[] { excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197496.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void ClearFormatting()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearFormatting");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196969.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAppendTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAppendTable");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839633.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void ToggleCharacterCode()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleCharacterCode");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821674.aspx </remarks>
        /// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PasteAndFormat", type);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837670.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838352.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void ShrinkDiscontiguousSelection()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShrinkDiscontiguousSelection");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838293.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void InsertStyleSeparator()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertStyleSeparator");
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
        /// <param name="editorID">optional object editorID</param>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToEditableRange(object editorID)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", typeof(NetOffice.WordApi.Range), editorID);
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Range GoToEditableRange()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", typeof(NetOffice.WordApi.Range));
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838493.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearParagraphStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearParagraphStyle");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191975.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearCharacterAllFormatting()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearCharacterAllFormatting");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841083.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearCharacterStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearCharacterStyle");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838672.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearCharacterDirectFormatting()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearCharacterDirectFormatting");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
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
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196419.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ReadingModeGrowFont()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReadingModeGrowFont");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196279.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ReadingModeShrinkFont()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReadingModeShrinkFont");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836876.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearParagraphAllFormatting()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearParagraphAllFormatting");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197502.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void ClearParagraphDirectFormatting()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearParagraphDirectFormatting");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195985.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void InsertNewPage()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertNewPage");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual void SortByHeadings()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SortByHeadings");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
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

