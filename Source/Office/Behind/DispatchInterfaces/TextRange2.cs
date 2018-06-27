using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// TextRange2
    /// </summary>
    [SyntaxBypass]
    public class TextRange2_ : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.TextRange2_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public TextRange2_() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Paragraphs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Paragraphs")]
        public virtual NetOffice.OfficeApi.TextRange2 Paragraphs(object start, object length)
        {
            return get_Paragraphs(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Paragraphs(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Paragraphs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Paragraphs")]
        public virtual NetOffice.OfficeApi.TextRange2 Paragraphs(object start)
        {
            return get_Paragraphs(start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Sentences(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Sentences
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Sentences")]
        public virtual NetOffice.OfficeApi.TextRange2 Sentences(object start, object length)
        {
            return get_Sentences(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Sentences(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Sentences
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Sentences")]
        public virtual NetOffice.OfficeApi.TextRange2 Sentences(object start)
        {
            return get_Sentences(start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Words(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Words
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Words")]
        public virtual NetOffice.OfficeApi.TextRange2 Words(object start, object length)
        {
            return get_Words(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Words(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Words
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Words")]
        public virtual NetOffice.OfficeApi.TextRange2 Words(object start)
        {
            return get_Words(start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Characters(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.OfficeApi.TextRange2 Characters(object start, object length)
        {
            return get_Characters(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Characters(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.OfficeApi.TextRange2 Characters(object start)
        {
            return get_Characters(start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Lines(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Lines
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Lines")]
        public virtual NetOffice.OfficeApi.TextRange2 Lines(object start, object length)
        {
            return get_Lines(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Lines(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Lines
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Lines")]
        public virtual NetOffice.OfficeApi.TextRange2 Lines(object start)
        {
            return get_Lines(start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Runs(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Runs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Runs")]
        public virtual NetOffice.OfficeApi.TextRange2 Runs(object start, object length)
        {
            return get_Runs(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_Runs(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Runs
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Runs")]
        public virtual NetOffice.OfficeApi.TextRange2 Runs(object start)
        {
            return get_Runs(start);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_MathZones(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", typeof(NetOffice.OfficeApi.TextRange2), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Alias for get_MathZones
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        /// <param name="length">optional Int32 length</param>
        [SupportByVersion("Office", 14, 15, 16), Redirect("get_MathZones")]
        public virtual NetOffice.OfficeApi.TextRange2 MathZones(object start, object length)
        {
            return get_MathZones(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional Int32 start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.TextRange2 get_MathZones(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", typeof(NetOffice.OfficeApi.TextRange2), start);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Alias for get_MathZones
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        /// <param name="start">optional Int32 start</param>
        [SupportByVersion("Office", 14, 15, 16), Redirect("get_MathZones")]
        public virtual NetOffice.OfficeApi.TextRange2 MathZones(object start)
        {
            return get_MathZones(start);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface TextRange2 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863528.aspx </remarks>
    public class TextRange2 : NetOffice.OfficeApi.Behind.TextRange2_, NetOffice.OfficeApi.TextRange2
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
                    _contractType = typeof(NetOffice.OfficeApi.TextRange2);
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
                    _type = typeof(TextRange2);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public TextRange2() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863807.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861203.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862210.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860549.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Paragraphs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Paragraphs", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860794.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Sentences
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Sentences", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864053.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Words
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Words", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863305.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Characters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Characters", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862044.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Lines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Lines", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861768.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Runs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "Runs", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862198.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ParagraphFormat2 ParagraphFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ParagraphFormat2>(this, "ParagraphFormat", typeof(NetOffice.OfficeApi.ParagraphFormat2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860218.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Font2 Font
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Font2>(this, "Font", typeof(NetOffice.OfficeApi.Font2));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861200.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Length
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Length");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861772.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Start
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Start");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863024.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BoundLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BoundLeft");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863847.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BoundTop
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BoundTop");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863508.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BoundWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BoundWidth");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860263.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Single BoundHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BoundHeight");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861366.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLanguageID LanguageID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "LanguageID");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageID", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860854.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 MathZones
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "MathZones", typeof(NetOffice.OfficeApi.TextRange2));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.TextRange2 this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Item", typeof(NetOffice.OfficeApi.TextRange2), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861091.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 TrimText()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "TrimText", typeof(NetOffice.OfficeApi.TextRange2));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862180.aspx </remarks>
        /// <param name="newText">optional string NewText = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertAfter(object newText)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertAfter", typeof(NetOffice.OfficeApi.TextRange2), newText);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862180.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertAfter()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertAfter", typeof(NetOffice.OfficeApi.TextRange2));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865495.aspx </remarks>
        /// <param name="newText">optional string NewText = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertBefore(object newText)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertBefore", typeof(NetOffice.OfficeApi.TextRange2), newText);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865495.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertBefore()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertBefore", typeof(NetOffice.OfficeApi.TextRange2));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862495.aspx </remarks>
        /// <param name="fontName">string fontName</param>
        /// <param name="charNumber">Int32 charNumber</param>
        /// <param name="unicode">optional NetOffice.OfficeApi.Enums.MsoTriState Unicode = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber, object unicode)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertSymbol", typeof(NetOffice.OfficeApi.TextRange2), fontName, charNumber, unicode);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862495.aspx </remarks>
        /// <param name="fontName">string fontName</param>
        /// <param name="charNumber">Int32 charNumber</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertSymbol(string fontName, Int32 charNumber)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertSymbol", typeof(NetOffice.OfficeApi.TextRange2), fontName, charNumber);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860564.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Select()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862117.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Cut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Cut");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863743.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862838.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863850.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Paste()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Paste", typeof(NetOffice.OfficeApi.TextRange2));
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862719.aspx </remarks>
        /// <param name="format">NetOffice.OfficeApi.Enums.MsoClipboardFormat format</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 PasteSpecial(NetOffice.OfficeApi.Enums.MsoClipboardFormat format)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "PasteSpecial", typeof(NetOffice.OfficeApi.TextRange2), format);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864574.aspx </remarks>
        /// <param name="type">NetOffice.OfficeApi.Enums.MsoTextChangeCase type</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChangeCase(NetOffice.OfficeApi.Enums.MsoTextChangeCase type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeCase", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861212.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AddPeriods()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddPeriods");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861820.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void RemovePeriods()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RemovePeriods");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        /// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase, object wholeWords)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", typeof(NetOffice.OfficeApi.TextRange2), findWhat, after, matchCase, wholeWords);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Find(string findWhat)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", typeof(NetOffice.OfficeApi.TextRange2), findWhat);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", typeof(NetOffice.OfficeApi.TextRange2), findWhat, after);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863750.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Find(string findWhat, object after, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Find", typeof(NetOffice.OfficeApi.TextRange2), findWhat, after, matchCase);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        /// <param name="wholeWords">optional NetOffice.OfficeApi.Enums.MsoTriState WholeWords = 0</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase, object wholeWords)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", typeof(NetOffice.OfficeApi.TextRange2), new object[] { findWhat, replaceWhat, after, matchCase, wholeWords });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", typeof(NetOffice.OfficeApi.TextRange2), findWhat, replaceWhat);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", typeof(NetOffice.OfficeApi.TextRange2), findWhat, replaceWhat, after);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864978.aspx </remarks>
        /// <param name="findWhat">string findWhat</param>
        /// <param name="replaceWhat">string replaceWhat</param>
        /// <param name="after">optional Int32 After = 0</param>
        /// <param name="matchCase">optional NetOffice.OfficeApi.Enums.MsoTriState MatchCase = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 Replace(string findWhat, string replaceWhat, object after, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "Replace", typeof(NetOffice.OfficeApi.TextRange2), findWhat, replaceWhat, after, matchCase);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865241.aspx </remarks>
        /// <param name="x1">Single x1</param>
        /// <param name="y1">Single y1</param>
        /// <param name="x2">Single x2</param>
        /// <param name="y2">Single y2</param>
        /// <param name="x3">Single x3</param>
        /// <param name="y3">Single y3</param>
        /// <param name="x4">Single x4</param>
        /// <param name="y4">Single y4</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void RotatedBounds(out Single x1, out Single y1, out Single x2, out Single y2, out Single x3, out Single y3, out Single x4, out Single y4)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true, true, true, true, true);
            x1 = 0;
            y1 = 0;
            x2 = 0;
            y2 = 0;
            x3 = 0;
            y3 = 0;
            x4 = 0;
            y4 = 0;
            object[] paramsArray = new object[] { x1, y1, x2, y2, x3, y3, x4, y4 };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "RotatedBounds", paramsArray, modifiers);

            x1 = (Single)paramsArray[0];
            y1 = (Single)paramsArray[1];
            x2 = (Single)paramsArray[2];
            y2 = (Single)paramsArray[3];
            x3 = (Single)paramsArray[4];
            y3 = (Single)paramsArray[5];
            x4 = (Single)paramsArray[6];
            y4 = (Single)paramsArray[7];
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861210.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void RtlRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RtlRun");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861750.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void LtrRun()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LtrRun");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        /// <param name="formula">optional string Formula = </param>
        /// <param name="position">optional Int32 Position = -1</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula, object position)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", typeof(NetOffice.OfficeApi.TextRange2), chartFieldType, formula, position);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", typeof(NetOffice.OfficeApi.TextRange2), chartFieldType);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227821.aspx </remarks>
        /// <param name="chartFieldType">NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType</param>
        /// <param name="formula">optional string Formula = </param>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public virtual NetOffice.OfficeApi.TextRange2 InsertChartField(NetOffice.OfficeApi.Enums.MsoChartFieldType chartFieldType, object formula)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.TextRange2>(this, "InsertChartField", typeof(NetOffice.OfficeApi.TextRange2), chartFieldType, formula);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.TextRange2>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.TextRange2>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.TextRange2>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.TextRange2>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.TextRange2> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.TextRange2 item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
