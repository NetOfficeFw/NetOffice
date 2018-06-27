using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi.Behind
{
    /// <summary>
    /// _Attachment
    /// </summary>
    [SyntaxBypass]
    public class _Attachment_ : NetOffice.OfficeApi.Behind.IAccessible, NetOffice.AccessApi._Attachment_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public _Attachment_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_FileName(object var)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileName", var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileName
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileName")]
        public virtual string FileName(object var)
        {
            return get_FileName(var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_FileType(object var)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileType", var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileType
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileType")]
        public virtual string FileType(object var)
        {
            return get_FileType(var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_FileURL(object var)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileURL", var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileURL
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx </remarks>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileURL")]
        public virtual string FileURL(object var)
        {
            return get_FileURL(var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_FileData(object var)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FileData", var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_FileData
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_FileData")]
        public virtual object FileData(object var)
        {
            return get_FileData(var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_PictureDisp(object var)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PictureDisp", var);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Alias for get_PictureDisp
        /// </summary>
        /// <param name="var">optional object var</param>
        [SupportByVersion("Access", 12, 14, 15, 16), Redirect("get_PictureDisp")]
        public virtual object PictureDisp(object var)
        {
            return get_PictureDisp(var);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface _Attachment 
    /// SupportByVersion Access, 12,14,15,16
    /// </summary>
    [SupportByVersion("Access", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _Attachment : _Attachment_, NetOffice.AccessApi._Attachment
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
                    _contractType = typeof(NetOffice.AccessApi._Attachment);
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
                    _type = typeof(_Attachment);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public _Attachment() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835342.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820946.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836659.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual object OldValue
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "OldValue");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192518.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Properties Properties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Properties>(this, "Properties", typeof(NetOffice.AccessApi.Properties));
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196010.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Children Controls
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Children>(this, "Controls", typeof(NetOffice.AccessApi.Children));
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [BaseResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.AccessApi._Hyperlink Hyperlink
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.AccessApi._Hyperlink>(this, "Hyperlink");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836727.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string EventProcPrefix
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EventProcPrefix");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EventProcPrefix", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string _Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "_Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "_Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836871.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte ControlType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "ControlType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844989.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte PictureSizeMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureSizeMode");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureSizeMode", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193807.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte PictureAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "PictureAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835074.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool PictureTiling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PictureTiling");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureTiling", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845492.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool Visible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Visible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195722.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte DisplayWhen
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DisplayWhen");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayWhen", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834511.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Left");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844867.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Top");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195738.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Width");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192489.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Height");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196447.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte BackStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "BackStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196184.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 BackColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193541.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte SpecialEffect
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "SpecialEffect");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpecialEffect", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820796.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte BorderStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "BorderStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821483.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte OldBorderStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "OldBorderStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OldBorderStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834465.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 BorderColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BorderColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835995.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte BorderWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "BorderWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual byte BorderLineStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "BorderLineStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderLineStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823119.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string ControlTipText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlTipText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlTipText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821449.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 HelpContextId
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HelpContextId");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HelpContextId", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837166.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 Section
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Section");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Section", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string ControlName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194772.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool IsVisible
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsVisible");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsVisible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193160.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string BeforeUpdate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeUpdate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeUpdate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194169.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string AfterUpdate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterUpdate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterUpdate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194485.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnEnter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnEnter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnEnter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194575.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnExit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnExit");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnExit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821070.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnDirty
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDirty");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDirty", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823022.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnChange
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnChange");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnChange", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820975.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnGotFocus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnGotFocus");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnGotFocus", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194110.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnLostFocus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLostFocus");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLostFocus", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822868.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnClick
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClick");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClick", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194550.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnDblClick
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDblClick");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDblClick", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195807.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnMouseDown
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseDown");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseDown", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193537.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnMouseMove
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseMove");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseMove", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844980.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnMouseUp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseUp");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseUp", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196490.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnKeyDown
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyDown");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyDown", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194243.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnKeyUp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyUp");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyUp", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198361.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnKeyPress
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyPress");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyPress", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196159.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string OnAttachmentCurrent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnAttachmentCurrent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnAttachmentCurrent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string BeforeUpdateMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BeforeUpdateMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BeforeUpdateMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string AfterUpdateMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AfterUpdateMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AfterUpdateMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnEnterMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnEnterMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnEnterMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnExitMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnExitMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnExitMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnDirtyMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDirtyMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDirtyMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnChangeMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnChangeMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnChangeMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnGotFocusMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnGotFocusMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnGotFocusMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnLostFocusMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnLostFocusMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnLostFocusMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnClickMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnClickMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnClickMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnDblClickMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnDblClickMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnDblClickMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnMouseDownMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseDownMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseDownMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnMouseMoveMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseMoveMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseMoveMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnMouseUpMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnMouseUpMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnMouseUpMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnKeyDownMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyDownMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyDownMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnKeyUpMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyUpMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyUpMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnKeyPressMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnKeyPressMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnKeyPressMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string OnAttachmentCurrentMacro
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OnAttachmentCurrentMacro");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OnAttachmentCurrentMacro", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822748.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string ShortcutMenuBar
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ShortcutMenuBar");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShortcutMenuBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835668.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool InSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InSelection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195530.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string Tag
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Tag");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Tag", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194767.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821403.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Enums.AcDisplayAs DisplayAs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcDisplayAs>(this, "DisplayAs");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayAs", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192274.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 AttachmentCount
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AttachmentCount");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197073.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 CurrentAttachment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CurrentAttachment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CurrentAttachment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193567.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string FileName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileName");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845243.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string FileType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileType");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195201.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string FileURL
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FileURL");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845195.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Enums.AcHorizontalAnchor HorizontalAnchor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcHorizontalAnchor>(this, "HorizontalAnchor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "HorizontalAnchor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822708.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Enums.AcVerticalAnchor VerticalAnchor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcVerticalAnchor>(this, "VerticalAnchor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "VerticalAnchor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845398.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual NetOffice.AccessApi.Enums.AcLayoutType Layout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcLayoutType>(this, "Layout");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195463.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 LeftPadding
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LeftPadding");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftPadding", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845619.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 TopPadding
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TopPadding");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopPadding", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821735.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 RightPadding
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "RightPadding");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightPadding", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822438.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 BottomPadding
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "BottomPadding");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BottomPadding", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845104.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineStyleLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineStyleLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineStyleLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822793.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineStyleTop
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineStyleTop");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineStyleTop", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823207.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineStyleRight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineStyleRight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineStyleRight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192328.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineStyleBottom
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineStyleBottom");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineStyleBottom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192080.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineWidthLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineWidthLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineWidthLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195855.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineWidthTop
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineWidthTop");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineWidthTop", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195871.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineWidthRight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineWidthRight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineWidthRight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193208.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte GridlineWidthBottom
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "GridlineWidthBottom");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineWidthBottom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836375.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 GridlineColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridlineColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197948.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string DefaultPicture
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultPicture");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultPicture", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835967.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int32 LayoutID
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LayoutID");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845041.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool AutoLabel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoLabel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoLabel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197385.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool AddColon
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AddColon");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddColon", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195238.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 LabelX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LabelX");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LabelX", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193492.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 LabelY
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "LabelY");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LabelY", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192309.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual byte LabelAlign
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "LabelAlign");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LabelAlign", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845588.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 ColumnWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ColumnWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835664.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 ColumnOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ColumnOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196054.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool ColumnHidden
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ColumnHidden");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ColumnHidden", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195499.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string ControlSource
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ControlSource");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ControlSource", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198248.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual string StatusBarText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StatusBarText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StatusBarText", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197633.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool TabStop
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "TabStop");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabStop", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822722.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual Int16 TabIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "TabIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TabIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834720.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool Enabled
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Enabled");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Enabled", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834789.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool Locked
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Locked");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Locked", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object FileData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "FileData");
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object PictureDisp
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "PictureDisp");
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823010.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Int32 BackThemeColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackThemeColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackThemeColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836938.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single BackTint
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BackTint");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackTint", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191790.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single BackShade
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BackShade");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BackShade", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820935.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Int32 BorderThemeColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BorderThemeColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderThemeColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196496.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single BorderTint
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BorderTint");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderTint", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192532.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single BorderShade
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "BorderShade");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BorderShade", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845642.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Int32 GridlineThemeColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridlineThemeColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineThemeColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822861.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single GridlineTint
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridlineTint");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineTint", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191829.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual Single GridlineShade
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "GridlineShade");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridlineShade", value);
            }
        }

        /// <summary>
        /// SupportByVersion Access 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196043.aspx </remarks>
        [SupportByVersion("Access", 14, 15, 16)]
        public virtual byte DefaultPictureType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBytePropertyGet(this, "DefaultPictureType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultPictureType", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198323.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void SizeToFit()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SizeToFit");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195435.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Requery()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Requery");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Goto()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Goto");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194081.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void SetFocus()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetFocus");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrExpr">string bstrExpr</param>
        /// <param name="ppsa">optional object[] ppsa</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual object _Evaluate(string bstrExpr, object[] ppsa)
        {
            object[] paramsArray = Invoker.ValidateParamsArray(bstrExpr, (object)ppsa);
            object returnItem = Invoker.MethodReturn(this, "_Evaluate", paramsArray);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                return newObject;
            }
            else
            {
                return returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrExpr">string bstrExpr</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual object _Evaluate(string bstrExpr)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", bstrExpr);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        /// <param name="width">optional object width</param>
        /// <param name="height">optional object height</param>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Move(object left, object top, object width, object height)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width, height);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Move(object left)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Move(object left, object top)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834482.aspx </remarks>
        /// <param name="left">object left</param>
        /// <param name="top">optional object top</param>
        /// <param name="width">optional object width</param>
        [CustomMethod]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Move(object left, object top, object width)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top, width);
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845316.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Forward()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Forward");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197753.aspx </remarks>
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual void Back()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Back");
        }

        /// <summary>
        /// SupportByVersion Access 12, 14, 15, 16
        /// </summary>
        /// <param name="dispid">Int32 dispid</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Access", 12, 14, 15, 16)]
        public virtual bool IsMemberSafe(Int32 dispid)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
        }

        #endregion

        #pragma warning restore
    }
}
