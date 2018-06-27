using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;
using NetOffice.CoreServices;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// DispatchInterface _Application
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _Application : COMObject, NetOffice.WordApi._Application
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
                    _contractType = typeof(NetOffice.WordApi._Application);
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
                    _type = typeof(_Application);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Application() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823254.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197825.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191758.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845178.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821628.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Documents Documents
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Documents>(this, "Documents", typeof(NetOffice.WordApi.Documents));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Windows Windows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Windows>(this, "Windows", typeof(NetOffice.WordApi.Windows));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837737.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document ActiveDocument
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "ActiveDocument", typeof(NetOffice.WordApi.Document));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845301.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Window ActiveWindow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Window>(this, "ActiveWindow", typeof(NetOffice.WordApi.Window));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838682.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Selection Selection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Selection>(this, "Selection", typeof(NetOffice.WordApi.Selection));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822917.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object WordBasic
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "WordBasic");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195679.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.RecentFiles RecentFiles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RecentFiles>(this, "RecentFiles", typeof(NetOffice.WordApi.RecentFiles));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845589.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Template NormalTemplate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Template>(this, "NormalTemplate", typeof(NetOffice.WordApi.Template));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822391.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.System System
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.System>(this, "System", typeof(NetOffice.WordApi.System));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845308.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.AutoCorrect AutoCorrect
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(this, "AutoCorrect", typeof(NetOffice.WordApi.AutoCorrect));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197817.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FontNames FontNames
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "FontNames", typeof(NetOffice.WordApi.FontNames));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196340.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FontNames LandscapeFontNames
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "LandscapeFontNames", typeof(NetOffice.WordApi.FontNames));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192201.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FontNames PortraitFontNames
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FontNames>(this, "PortraitFontNames", typeof(NetOffice.WordApi.FontNames));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840701.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Languages Languages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Languages>(this, "Languages", typeof(NetOffice.WordApi.Languages));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Assistant Assistant
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Assistant>(this, "Assistant", typeof(NetOffice.OfficeApi.Assistant));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821300.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Browser Browser
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Browser>(this, "Browser", typeof(NetOffice.WordApi.Browser));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823259.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.FileConverters FileConverters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FileConverters>(this, "FileConverters", typeof(NetOffice.WordApi.FileConverters));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821659.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.MailingLabel MailingLabel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailingLabel>(this, "MailingLabel", typeof(NetOffice.WordApi.MailingLabel));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191745.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Dialogs Dialogs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dialogs>(this, "Dialogs", typeof(NetOffice.WordApi.Dialogs));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838479.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.CaptionLabels CaptionLabels
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CaptionLabels>(this, "CaptionLabels", typeof(NetOffice.WordApi.CaptionLabels));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198063.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.AutoCaptions AutoCaptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCaptions>(this, "AutoCaptions", typeof(NetOffice.WordApi.AutoCaptions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.AddIns AddIns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AddIns>(this, "AddIns", typeof(NetOffice.WordApi.AddIns));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839544.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
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
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821519.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Version
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197438.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ScreenUpdating
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ScreenUpdating");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ScreenUpdating", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198164.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool PrintPreview
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PrintPreview");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PrintPreview", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839740.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Tasks Tasks
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tasks>(this, "Tasks", typeof(NetOffice.WordApi.Tasks));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayStatusBar
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayStatusBar");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayStatusBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836086.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool SpecialMode
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "SpecialMode");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839688.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 UsableWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UsableWidth");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834606.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 UsableHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "UsableHeight");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192165.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MathCoprocessorAvailable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MathCoprocessorAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192426.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MouseAvailable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MouseAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_International(NetOffice.WordApi.Enums.WdInternationalIndex index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "International", index);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_International
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_International")]
        public virtual object International(NetOffice.WordApi.Enums.WdInternationalIndex index)
        {
            return get_International(index);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839495.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Build
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Build");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820850.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CapsLock
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CapsLock");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845392.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool NumLock
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NumLock");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834599.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string UserName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserName", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844813.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string UserInitials
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserInitials");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserInitials", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193411.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string UserAddress
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "UserAddress");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UserAddress", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835128.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object MacroContainer
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MacroContainer");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838964.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayRecentFiles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayRecentFiles");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayRecentFiles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845623.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.CommandBars CommandBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.CommandBars>(this, "CommandBars", typeof(NetOffice.OfficeApi.CommandBars));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word, object languageID)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", typeof(NetOffice.WordApi.SynonymInfo), word, languageID);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        public virtual NetOffice.WordApi.SynonymInfo SynonymInfo(string word, object languageID)
        {
            return get_SynonymInfo(word, languageID);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", typeof(NetOffice.WordApi.SynonymInfo), word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        public virtual NetOffice.WordApi.SynonymInfo SynonymInfo(string word)
        {
            return get_SynonymInfo(word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197234.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", typeof(NetOffice.VBIDEApi.VBE));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839412.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string DefaultSaveFormat
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultSaveFormat");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultSaveFormat", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821102.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.ListGalleries ListGalleries
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListGalleries>(this, "ListGalleries", typeof(NetOffice.WordApi.ListGalleries));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821995.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ActivePrinter
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ActivePrinter");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ActivePrinter", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821925.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Templates Templates
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Templates>(this, "Templates", typeof(NetOffice.WordApi.Templates));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822548.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object CustomizationContext
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "CustomizationContext");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "CustomizationContext", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197596.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.KeyBindings KeyBindings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBindings>(this, "KeyBindings", typeof(NetOffice.WordApi.KeyBindings));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        /// <param name="commandParameter">optional object commandParameter</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(this, "KeysBoundTo", typeof(NetOffice.WordApi.KeysBoundTo), keyCategory, command, commandParameter);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        /// <param name="commandParameter">optional object commandParameter</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        public virtual NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter)
        {
            return get_KeysBoundTo(keyCategory, command, commandParameter);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeysBoundTo>(this, "KeysBoundTo", typeof(NetOffice.WordApi.KeysBoundTo), keyCategory, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        public virtual NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command)
        {
            return get_KeysBoundTo(keyCategory, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode, object keyCode2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(this, "FindKey", typeof(NetOffice.WordApi.KeyBinding), keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        public virtual NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode, object keyCode2)
        {
            return get_FindKey(keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.KeyBinding>(this, "FindKey", typeof(NetOffice.WordApi.KeyBinding), keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        public virtual NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode)
        {
            return get_FindKey(keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196028.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Caption
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192216.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string Path
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192367.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayScrollBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayScrollBars");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayScrollBars", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string StartupPath
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StartupPath");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StartupPath", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835146.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BackgroundSavingStatus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackgroundSavingStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820962.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BackgroundPrintingStatus
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BackgroundPrintingStatus");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839318.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Left");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837463.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Top");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836284.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Width");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Width", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845159.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Height");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Height", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836388.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdWindowState WindowState
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdWindowState>(this, "WindowState");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "WindowState", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192152.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayAutoCompleteTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayAutoCompleteTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayAutoCompleteTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822542.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Options Options
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Options>(this, "Options", typeof(NetOffice.WordApi.Options));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192373.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdAlertLevel DisplayAlerts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAlertLevel>(this, "DisplayAlerts");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayAlerts", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191957.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Dictionaries CustomDictionaries
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Dictionaries>(this, "CustomDictionaries", typeof(NetOffice.WordApi.Dictionaries));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192616.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string PathSeparator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "PathSeparator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845291.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string StatusBar
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "StatusBar");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StatusBar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192800.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MAPIAvailable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MAPIAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845182.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool DisplayScreenTips
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayScreenTips");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayScreenTips", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839294.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdEnableCancelKey EnableCancelKey
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEnableCancelKey>(this, "EnableCancelKey");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EnableCancelKey", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197424.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool UserControl
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UserControl");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.FileSearch FileSearch
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileSearch>(this, "FileSearch", typeof(NetOffice.OfficeApi.FileSearch));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838972.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.WdMailSystem MailSystem
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdMailSystem>(this, "MailSystem");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string DefaultTableSeparator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultTableSeparator");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultTableSeparator", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowVisualBasicEditor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowVisualBasicEditor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowVisualBasicEditor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839549.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string BrowseExtraFileTypes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BrowseExtraFileTypes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BrowseExtraFileTypes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool get_IsObjectValid(object _object)
        {
            return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsObjectValid", _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_IsObjectValid
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_IsObjectValid")]
        public virtual bool IsObjectValid(object _object)
        {
            return get_IsObjectValid(_object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194713.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.HangulHanjaConversionDictionaries HangulHanjaDictionaries
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HangulHanjaConversionDictionaries>(this, "HangulHanjaDictionaries", typeof(NetOffice.WordApi.HangulHanjaConversionDictionaries));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.MailMessage MailMessage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.MailMessage>(this, "MailMessage", typeof(NetOffice.WordApi.MailMessage));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840871.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool FocusInMailHeader
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "FocusInMailHeader");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192588.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.EmailOptions EmailOptions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EmailOptions>(this, "EmailOptions", typeof(NetOffice.WordApi.EmailOptions));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836711.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLanguageID Language
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLanguageID>(this, "Language");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192831.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.COMAddIns COMAddIns
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.COMAddIns>(this, "COMAddIns", typeof(NetOffice.OfficeApi.COMAddIns));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192428.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckLanguage
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CheckLanguage");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CheckLanguage", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.LanguageSettings LanguageSettings
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LanguageSettings>(this, "LanguageSettings", typeof(NetOffice.OfficeApi.LanguageSettings));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool Dummy1
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Dummy1");
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.AnswerWizard AnswerWizard
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.AnswerWizard>(this, "AnswerWizard", typeof(NetOffice.OfficeApi.AnswerWizard));
            }
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195192.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFeatureInstall>(this, "FeatureInstall");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FeatureInstall", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192776.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutomationSecurity>(this, "AutomationSecurity");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AutomationSecurity", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FileDialog>(this, "FileDialog", typeof(NetOffice.OfficeApi.FileDialog), fileDialogType);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16), Redirect("get_FileDialog")]
        public virtual NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType)
        {
            return get_FileDialog(fileDialogType);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193382.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual string EmailTemplate
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EmailTemplate");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "EmailTemplate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowWindowsInTaskbar
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowWindowsInTaskbar");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowWindowsInTaskbar", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193065.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.NewFile NewDocument
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.NewFile>(this, "NewDocument", typeof(NetOffice.OfficeApi.NewFile));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840052.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowStartupDialog
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowStartupDialog");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowStartupDialog", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192177.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.AutoCorrect AutoCorrectEmail
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.AutoCorrect>(this, "AutoCorrectEmail", typeof(NetOffice.WordApi.AutoCorrect));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845341.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.TaskPanes TaskPanes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TaskPanes>(this, "TaskPanes", typeof(NetOffice.WordApi.TaskPanes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835491.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool DefaultLegalBlackline
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DefaultLegalBlackline");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultLegalBlackline", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SmartTagRecognizers SmartTagRecognizers
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagRecognizers>(this, "SmartTagRecognizers", typeof(NetOffice.WordApi.SmartTagRecognizers));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SmartTagTypes SmartTagTypes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTagTypes>(this, "SmartTagTypes", typeof(NetOffice.WordApi.SmartTagTypes));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839771.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.XMLNamespaces XMLNamespaces
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNamespaces>(this, "XMLNamespaces", typeof(NetOffice.WordApi.XMLNamespaces));
            }
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196679.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual bool ArbitraryXMLSupportAvailable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ArbitraryXMLSupportAvailable");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string BuildFull
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BuildFull");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string BuildFeatureCrew
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BuildFeatureCrew");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192405.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Bibliography Bibliography
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bibliography>(this, "Bibliography", typeof(NetOffice.WordApi.Bibliography));
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191727.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual bool ShowStylePreviews
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowStylePreviews");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowStylePreviews", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845435.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual bool RestrictLinkedStyles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RestrictLinkedStyles");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RestrictLinkedStyles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837322.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.OMathAutoCorrect OMathAutoCorrect
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMathAutoCorrect>(this, "OMathAutoCorrect", typeof(NetOffice.WordApi.OMathAutoCorrect));
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836074.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual bool DisplayDocumentInformationPanel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayDocumentInformationPanel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayDocumentInformationPanel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197133.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IAssistance Assistance
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IAssistance>(this, "Assistance", typeof(NetOffice.OfficeApi.IAssistance));
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192620.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual bool OpenAttachmentsInFullScreen
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "OpenAttachmentsInFullScreen");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OpenAttachmentsInFullScreen", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836063.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual Int32 ActiveEncryptionSession
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ActiveEncryptionSession");
            }
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194203.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual bool DontResetInsertionPointProperties
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DontResetInsertionPointProperties");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DontResetInsertionPointProperties", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839192.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtLayouts>(this, "SmartArtLayouts", typeof(NetOffice.OfficeApi.SmartArtLayouts));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194982.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtQuickStyles>(this, "SmartArtQuickStyles", typeof(NetOffice.OfficeApi.SmartArtQuickStyles));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839505.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.SmartArtColors SmartArtColors
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SmartArtColors>(this, "SmartArtColors", typeof(NetOffice.OfficeApi.SmartArtColors));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838675.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.UndoRecord UndoRecord
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.UndoRecord>(this, "UndoRecord", typeof(NetOffice.WordApi.UndoRecord));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191978.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.PickerDialog PickerDialog
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PickerDialog>(this, "PickerDialog", typeof(NetOffice.OfficeApi.PickerDialog));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839925.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.ProtectedViewWindows ProtectedViewWindows
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindows>(this, "ProtectedViewWindows", typeof(NetOffice.WordApi.ProtectedViewWindows));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192773.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.ProtectedViewWindow ActiveProtectedViewWindow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProtectedViewWindow>(this, "ActiveProtectedViewWindow", typeof(NetOffice.WordApi.ProtectedViewWindow));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845787.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual bool IsSandboxed
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsSandboxed");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193078.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoFileValidationMode>(this, "FileValidation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FileValidation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232091.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual bool ChartDataPointTrack
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ChartDataPointTrack");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ChartDataPointTrack", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232207.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual bool ShowAnimation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowAnimation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowAnimation", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="originalFormat">optional object originalFormat</param>
        /// <param name="routeDocument">optional object routeDocument</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Quit(object saveChanges, object originalFormat, object routeDocument)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Quit", saveChanges, originalFormat, routeDocument);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Quit()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Quit(object saveChanges)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Quit", saveChanges);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="originalFormat">optional object originalFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Quit(object saveChanges, object originalFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Quit", saveChanges, originalFormat);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ScreenRefresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ScreenRefresh");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOutOld", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839803.aspx </remarks>
        /// <param name="name">string name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void LookupNameProperties(string name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LookupNameProperties", name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192415.aspx </remarks>
        /// <param name="unavailableFont">string unavailableFont</param>
        /// <param name="substituteFont">string substituteFont</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SubstituteFont(string unavailableFont, string substituteFont)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SubstituteFont", unavailableFont, substituteFont);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        /// <param name="times">optional object times</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Repeat(object times)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Repeat", times);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Repeat()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Repeat");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845561.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DDEExecute(Int32 channel, string command)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DDEExecute", channel, command);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837295.aspx </remarks>
        /// <param name="app">string app</param>
        /// <param name="topic">string topic</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DDEInitiate(string app, string topic)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DDEInitiate", app, topic);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837201.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        /// <param name="data">string data</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DDEPoke(Int32 channel, string item, string data)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DDEPoke", channel, item, data);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837546.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string DDERequest(Int32 channel, string item)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "DDERequest", channel, item);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837904.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DDETerminate(Int32 channel)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DDETerminate", channel);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192053.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DDETerminateAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DDETerminateAll");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "BuildKeyCode", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string KeyString(Int32 keyCode, object keyCode2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "KeyString", keyCode, keyCode2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string KeyString(Int32 keyCode)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "KeyString", keyCode);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835492.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="destination">string destination</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OrganizerCopy(string source, string destination, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OrganizerCopy", source, destination, name, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194744.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OrganizerDelete(string source, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OrganizerDelete", source, name, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836140.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OrganizerRename(string source, string name, string newName, NetOffice.WordApi.Enums.WdOrganizerObject _object)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OrganizerRename", source, name, newName, _object);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823266.aspx </remarks>
        /// <param name="tagID">String[] tagID</param>
        /// <param name="value">String[] value</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AddAddress(String[] tagID, String[] value)
        {
            object[] paramsArray = Invoker.ValidateParamsArray((object)tagID, (object)value);
            Invoker.Method(this, "AddAddress", paramsArray); ;
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        /// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
        /// <param name="updateRecentAddresses">optional object updateRecentAddresses</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice, object updateRecentAddresses)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice, updateRecentAddresses });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties, useAutoText);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", name, addressProperties, useAutoText, displaySelectDialog);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        /// <param name="displaySelectDialog">optional object displaySelectDialog</param>
        /// <param name="selectDialog">optional object selectDialog</param>
        /// <param name="checkNamesDialog">optional object checkNamesDialog</param>
        /// <param name="recentAddressesChoice">optional object recentAddressesChoice</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetAddress", new object[] { name, addressProperties, useAutoText, displaySelectDialog, selectDialog, checkNamesDialog, recentAddressesChoice });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194798.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckGrammar(string _string)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckGrammar", _string);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
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
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", word, customDictionary, ignoreUppercase, mainDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        /// <param name="customDictionary5">optional object customDictionary5</param>
        /// <param name="customDictionary6">optional object customDictionary6</param>
        /// <param name="customDictionary7">optional object customDictionary7</param>
        /// <param name="customDictionary8">optional object customDictionary8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
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
        public virtual bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CheckSpelling", new object[] { word, customDictionary, ignoreUppercase, mainDictionary, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822681.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ResetIgnoreAll()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetIgnoreAll");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), word);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), word, customDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), word, customDictionary, ignoreUppercase);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), word, customDictionary, ignoreUppercase, mainDictionary);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        /// <param name="mainDictionary">optional object mainDictionary</param>
        /// <param name="suggestionMode">optional object suggestionMode</param>
        /// <param name="customDictionary2">optional object customDictionary2</param>
        /// <param name="customDictionary3">optional object customDictionary3</param>
        /// <param name="customDictionary4">optional object customDictionary4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
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
        public virtual NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", typeof(NetOffice.WordApi.SpellingSuggestions), new object[] { word, customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838545.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void GoBack()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "GoBack");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841057.aspx </remarks>
        /// <param name="helpType">object helpType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Help(object helpType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Help", helpType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194337.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void AutomaticChange()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutomaticChange");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ShowMe()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowMe");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821932.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void HelpTool()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "HelpTool");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845336.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Window NewWindow()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Window>(this, "NewWindow", typeof(NetOffice.WordApi.Window));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194509.aspx </remarks>
        /// <param name="listAllCommands">bool listAllCommands</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ListCommands(bool listAllCommands)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ListCommands", listAllCommands);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834517.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ShowClipboard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowClipboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        /// <param name="tolerance">optional object tolerance</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OnTime(object when, string name, object tolerance)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OnTime", when, name, tolerance);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void OnTime(object when, string name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "OnTime", when, name);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837154.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void NextLetter()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "NextLetter");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        /// <param name="userPassword">optional object userPassword</param>
        /// <param name="volumePassword">optional object volumePassword</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int16 MountVolume(string zone, string server, string volume, object user, object userPassword, object volumePassword)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "MountVolume", new object[] { zone, server, volume, user, userPassword, volumePassword });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int16 MountVolume(string zone, string server, string volume)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "MountVolume", zone, server, volume);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int16 MountVolume(string zone, string server, string volume, object user)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "MountVolume", zone, server, volume, user);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        /// <param name="userPassword">optional object userPassword</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int16 MountVolume(string zone, string server, string volume, object user, object userPassword)
        {
            return InvokerService.InvokeInternal.ExecuteInt16MethodGet(this, "MountVolume", new object[] { zone, server, volume, user, userPassword });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844818.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string CleanString(string _string)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "CleanString", _string);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SendFax()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SendFax");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835219.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ChangeFileOpenDirectory(string path)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChangeFileOpenDirectory", path);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macroName">string macroName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void RunOld(string macroName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RunOld", macroName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void GoForward()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "GoForward");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844914.aspx </remarks>
        /// <param name="left">Int32 left</param>
        /// <param name="top">Int32 top</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Move(Int32 left, Int32 top)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Move", left, top);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197452.aspx </remarks>
        /// <param name="width">Int32 width</param>
        /// <param name="height">Int32 height</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Resize(Int32 width, Int32 height)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Resize", width, height);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197549.aspx </remarks>
        /// <param name="inches">Single inches</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single InchesToPoints(Single inches)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "InchesToPoints", inches);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838159.aspx </remarks>
        /// <param name="centimeters">Single centimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single CentimetersToPoints(Single centimeters)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "CentimetersToPoints", centimeters);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845767.aspx </remarks>
        /// <param name="millimeters">Single millimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single MillimetersToPoints(Single millimeters)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "MillimetersToPoints", millimeters);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840225.aspx </remarks>
        /// <param name="picas">Single picas</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PicasToPoints(Single picas)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PicasToPoints", picas);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840343.aspx </remarks>
        /// <param name="lines">Single lines</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single LinesToPoints(Single lines)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "LinesToPoints", lines);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838268.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToInches(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToInches", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195052.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToCentimeters(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToCentimeters", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836929.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToMillimeters(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToMillimeters", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193434.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToPicas(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToPicas", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822110.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToLines(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToLines", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Activate()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Activate");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToPixels(Single points, object fVertical)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToPixels", points, fVertical);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PointsToPixels(Single points)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PointsToPixels", points);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PixelsToPoints(Single pixels, object fVertical)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PixelsToPoints", pixels, fVertical);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single PixelsToPoints(Single pixels)
        {
            return InvokerService.InvokeInternal.ExecuteSingleMethodGet(this, "PixelsToPoints", pixels);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845662.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void KeyboardLatin()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "KeyboardLatin");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196621.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void KeyboardBidi()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "KeyboardBidi");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835971.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void ToggleKeyboard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ToggleKeyboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        /// <param name="langId">optional Int32 LangId = 0</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Keyboard(object langId)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Keyboard", langId);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Keyboard()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Keyboard");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193728.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string ProductCode()
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ProductCode");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840160.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.DefaultWebOptions DefaultWebOptions()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.DefaultWebOptions>(this, "DefaultWebOptions", typeof(NetOffice.WordApi.DefaultWebOptions));
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="range">object range</param>
        /// <param name="cid">object cid</param>
        /// <param name="piCSE">object piCSE</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void DiscussionSupport(object range, object cid, object piCSE)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DiscussionSupport", range, cid, piCSE);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821531.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void SetDefaultTheme(string name, NetOffice.WordApi.Enums.WdDocumentMedium documentType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultTheme", name, documentType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834585.aspx </remarks>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual string GetDefaultTheme(NetOffice.WordApi.Enums.WdDocumentMedium documentType)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetDefaultTheme", documentType);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        /// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        /// <param name="varg29">optional object varg29</param>
        /// <param name="varg30">optional object varg30</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29, object varg30)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", macroName);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", macroName, varg1);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", macroName, varg1, varg2);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", macroName, varg1, varg2, varg3);
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28 });
        }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        /// <param name="varg3">optional object varg3</param>
        /// <param name="varg4">optional object varg4</param>
        /// <param name="varg5">optional object varg5</param>
        /// <param name="varg6">optional object varg6</param>
        /// <param name="varg7">optional object varg7</param>
        /// <param name="varg8">optional object varg8</param>
        /// <param name="varg9">optional object varg9</param>
        /// <param name="varg10">optional object varg10</param>
        /// <param name="varg11">optional object varg11</param>
        /// <param name="varg12">optional object varg12</param>
        /// <param name="varg13">optional object varg13</param>
        /// <param name="varg14">optional object varg14</param>
        /// <param name="varg15">optional object varg15</param>
        /// <param name="varg16">optional object varg16</param>
        /// <param name="varg17">optional object varg17</param>
        /// <param name="varg18">optional object varg18</param>
        /// <param name="varg19">optional object varg19</param>
        /// <param name="varg20">optional object varg20</param>
        /// <param name="varg21">optional object varg21</param>
        /// <param name="varg22">optional object varg22</param>
        /// <param name="varg23">optional object varg23</param>
        /// <param name="varg24">optional object varg24</param>
        /// <param name="varg25">optional object varg25</param>
        /// <param name="varg26">optional object varg26</param>
        /// <param name="varg27">optional object varg27</param>
        /// <param name="varg28">optional object varg28</param>
        /// <param name="varg29">optional object varg29</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Run", new object[] { macroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29 });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        /// <param name="printZoomPaperHeight">optional object printZoomPaperHeight</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth, printZoomPaperHeight });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000");
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", background, append, range, outputFileName);
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        /// <param name="outputFileName">optional object outputFileName</param>
        /// <param name="from">optional object from</param>
        /// <param name="to">optional object to</param>
        /// <param name="item">optional object item</param>
        /// <param name="copies">optional object copies</param>
        /// <param name="pages">optional object pages</param>
        /// <param name="pageType">optional object pageType</param>
        /// <param name="printToFile">optional object printToFile</param>
        /// <param name="collate">optional object collate</param>
        /// <param name="fileName">optional object fileName</param>
        /// <param name="activePrinterMacGX">optional object activePrinterMacGX</param>
        /// <param name="manualDuplexPrint">optional object manualDuplexPrint</param>
        /// <param name="printZoomColumn">optional object printZoomColumn</param>
        /// <param name="printZoomRow">optional object printZoomRow</param>
        /// <param name="printZoomPaperWidth">optional object printZoomPaperWidth</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut2000", new object[] { background, append, range, outputFileName, from, to, item, copies, pages, pageType, printToFile, collate, fileName, activePrinterMacGX, manualDuplexPrint, printZoomColumn, printZoomRow, printZoomPaperWidth });
        }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        public virtual bool Dummy2()
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Dummy2");
        }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838158.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        public virtual void PutFocusInMailHeader()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PutFocusInMailHeader");
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840673.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual void LoadMasterList(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "LoadMasterList", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        /// <param name="ignoreAllComparisonWarnings">optional bool IgnoreAllComparisonWarnings = false</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor, object ignoreAllComparisonWarnings)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor, ignoreAllComparisonWarnings });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument, destination);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument, destination, granularity);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "CompareDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, revisedAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        /// <param name="formatFrom">optional NetOffice.WordApi.Enums.WdMergeFormatFrom FormatFrom = 2</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor, object formatFrom)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor, formatFrom });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument, destination);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), originalDocument, revisedDocument, destination, granularity);
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        /// <param name="granularity">optional NetOffice.WordApi.Enums.WdGranularity Granularity = 1</param>
        /// <param name="compareFormatting">optional bool CompareFormatting = true</param>
        /// <param name="compareCaseChanges">optional bool CompareCaseChanges = true</param>
        /// <param name="compareWhitespace">optional bool CompareWhitespace = true</param>
        /// <param name="compareTables">optional bool CompareTables = true</param>
        /// <param name="compareHeaders">optional bool CompareHeaders = true</param>
        /// <param name="compareFootnotes">optional bool CompareFootnotes = true</param>
        /// <param name="compareTextboxes">optional bool CompareTextboxes = true</param>
        /// <param name="compareFields">optional bool CompareFields = true</param>
        /// <param name="compareComments">optional bool CompareComments = true</param>
        /// <param name="compareMoves">optional bool CompareMoves = true</param>
        /// <param name="originalAuthor">optional string OriginalAuthor = </param>
        /// <param name="revisedAuthor">optional string RevisedAuthor = </param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        public virtual NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "MergeDocuments", typeof(NetOffice.WordApi.Document), new object[] { originalDocument, revisedDocument, destination, granularity, compareFormatting, compareCaseChanges, compareWhitespace, compareTables, compareHeaders, compareFootnotes, compareTextboxes, compareFields, compareComments, compareMoves, originalAuthor, revisedAuthor });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="localDocument">NetOffice.WordApi.Document localDocument</param>
        /// <param name="serverDocument">NetOffice.WordApi.Document serverDocument</param>
        /// <param name="baseDocument">NetOffice.WordApi.Document baseDocument</param>
        /// <param name="favorSource">bool favorSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ThreeWayMerge(NetOffice.WordApi.Document localDocument, NetOffice.WordApi.Document serverDocument, NetOffice.WordApi.Document baseDocument, bool favorSource)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ThreeWayMerge", localDocument, serverDocument, baseDocument, favorSource);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Dummy4()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Dummy4");
        }

        #endregion

        #pragma warning restore
    }
}
