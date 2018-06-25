using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
    /// <summary>
    /// DispatchInterface _Application
    /// SupportByVersion Word, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("00020970-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.WordApi.Application))]
    public interface _Application : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823254.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197825.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Creator { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191758.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845178.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821628.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Documents Documents { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Windows Windows { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837737.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Document ActiveDocument { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845301.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Window ActiveWindow { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838682.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Selection Selection { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822917.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object WordBasic { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195679.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.RecentFiles RecentFiles { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845589.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Template NormalTemplate { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822391.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.System System { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845308.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.AutoCorrect AutoCorrect { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197817.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FontNames FontNames { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196340.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FontNames LandscapeFontNames { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192201.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FontNames PortraitFontNames { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840701.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Languages Languages { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Assistant Assistant { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821300.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Browser Browser { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823259.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.FileConverters FileConverters { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821659.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.MailingLabel MailingLabel { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191745.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Dialogs Dialogs { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838479.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.CaptionLabels CaptionLabels { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198063.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.AutoCaptions AutoCaptions { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.AddIns AddIns { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839544.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool Visible { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821519.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Version { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197438.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool ScreenUpdating { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198164.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool PrintPreview { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839740.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Tasks Tasks { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayStatusBar { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836086.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool SpecialMode { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839688.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UsableWidth { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834606.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 UsableHeight { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192165.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool MathCoprocessorAvailable { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192426.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool MouseAvailable { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        object get_International(NetOffice.WordApi.Enums.WdInternationalIndex index);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_International
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823245.aspx </remarks>
        /// <param name="index">NetOffice.WordApi.Enums.WdInternationalIndex index</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_International")]
        object International(NetOffice.WordApi.Enums.WdInternationalIndex index);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839495.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Build { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820850.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CapsLock { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845392.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool NumLock { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834599.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string UserName { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844813.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string UserInitials { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193411.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string UserAddress { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835128.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object MacroContainer { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838964.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayRecentFiles { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845623.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.CommandBars CommandBars { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word, object languageID);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="languageID">optional object languageID</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        NetOffice.WordApi.SynonymInfo SynonymInfo(string word, object languageID);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.SynonymInfo get_SynonymInfo(string word);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_SynonymInfo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821393.aspx </remarks>
        /// <param name="word">string word</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_SynonymInfo")]
        NetOffice.WordApi.SynonymInfo SynonymInfo(string word);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197234.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.VBIDEApi.VBE VBE { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839412.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string DefaultSaveFormat { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821102.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.ListGalleries ListGalleries { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821995.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string ActivePrinter { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821925.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Templates Templates { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822548.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object CustomizationContext { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197596.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.KeyBindings KeyBindings { get; }

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
        NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        /// <param name="commandParameter">optional object commandParameter</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command, object commandParameter);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.KeysBoundTo get_KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_KeysBoundTo
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196068.aspx </remarks>
        /// <param name="keyCategory">NetOffice.WordApi.Enums.WdKeyCategory keyCategory</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_KeysBoundTo")]
        NetOffice.WordApi.KeysBoundTo KeysBoundTo(NetOffice.WordApi.Enums.WdKeyCategory keyCategory, string command);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode, object keyCode2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode, object keyCode2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.WordApi.KeyBinding get_FindKey(Int32 keyCode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_FindKey
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840614.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_FindKey")]
        NetOffice.WordApi.KeyBinding FindKey(Int32 keyCode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196028.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Caption { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192216.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string Path { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192367.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayScrollBars { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string StartupPath { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835146.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BackgroundSavingStatus { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820962.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BackgroundPrintingStatus { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839318.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Left { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837463.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Top { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836284.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Width { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845159.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Height { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836388.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdWindowState WindowState { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192152.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayAutoCompleteTips { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822542.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Options Options { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192373.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdAlertLevel DisplayAlerts { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191957.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Dictionaries CustomDictionaries { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192616.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string PathSeparator { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845291.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string StatusBar { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192800.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool MAPIAvailable { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845182.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool DisplayScreenTips { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839294.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdEnableCancelKey EnableCancelKey { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197424.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool UserControl { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.FileSearch FileSearch { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838972.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Enums.WdMailSystem MailSystem { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839937.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string DefaultTableSeparator { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool ShowVisualBasicEditor { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839549.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string BrowseExtraFileTypes { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool get_IsObjectValid(object _object);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Alias for get_IsObjectValid
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834540.aspx </remarks>
        /// <param name="_object">object object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16), Redirect("get_IsObjectValid")]
        bool IsObjectValid(object _object);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194713.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.HangulHanjaConversionDictionaries HangulHanjaDictionaries { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821986.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.MailMessage MailMessage { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840871.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool FocusInMailHeader { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192588.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.EmailOptions EmailOptions { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836711.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoLanguageID Language { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192831.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.COMAddIns COMAddIns { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192428.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CheckLanguage { get; set; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197161.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.LanguageSettings LanguageSettings { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        bool Dummy1 { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.AnswerWizard AnswerWizard { get; }

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195192.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoFeatureInstall FeatureInstall { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192776.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoAutomationSecurity AutomationSecurity { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        NetOffice.OfficeApi.FileDialog get_FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Alias for get_FileDialog
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840210.aspx </remarks>
        /// <param name="fileDialogType">NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType</param>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16), Redirect("get_FileDialog")]
        NetOffice.OfficeApi.FileDialog FileDialog(NetOffice.OfficeApi.Enums.MsoFileDialogType fileDialogType);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193382.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        string EmailTemplate { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool ShowWindowsInTaskbar { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193065.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.OfficeApi.NewFile NewDocument { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840052.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool ShowStartupDialog { get; set; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192177.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.AutoCorrect AutoCorrectEmail { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845341.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.TaskPanes TaskPanes { get; }

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835491.aspx </remarks>
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool DefaultLegalBlackline { get; set; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SmartTagRecognizers SmartTagRecognizers { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SmartTagTypes SmartTagTypes { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839771.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        NetOffice.WordApi.XMLNamespaces XMLNamespaces { get; }

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196679.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        bool ArbitraryXMLSupportAvailable { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string BuildFull { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        string BuildFeatureCrew { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192405.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.Bibliography Bibliography { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191727.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        bool ShowStylePreviews { get; set; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845435.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        bool RestrictLinkedStyles { get; set; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837322.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.OMathAutoCorrect OMathAutoCorrect { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836074.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        bool DisplayDocumentInformationPanel { get; set; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197133.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.OfficeApi.IAssistance Assistance { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192620.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        bool OpenAttachmentsInFullScreen { get; set; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836063.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        Int32 ActiveEncryptionSession { get; }

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194203.aspx </remarks>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        bool DontResetInsertionPointProperties { get; set; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839192.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtLayouts SmartArtLayouts { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194982.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtQuickStyles SmartArtQuickStyles { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839505.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.SmartArtColors SmartArtColors { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838675.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.UndoRecord UndoRecord { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191978.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.PickerDialog PickerDialog { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839925.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ProtectedViewWindows ProtectedViewWindows { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192773.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.WordApi.ProtectedViewWindow ActiveProtectedViewWindow { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845787.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        bool IsSandboxed { get; }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193078.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoFileValidationMode FileValidation { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232091.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        bool ChartDataPointTrack { get; set; }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232207.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        bool ShowAnimation { get; set; }

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
        void Quit(object saveChanges, object originalFormat, object routeDocument);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Quit();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Quit(object saveChanges);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844895.aspx </remarks>
        /// <param name="saveChanges">optional object saveChanges</param>
        /// <param name="originalFormat">optional object originalFormat</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Quit(object saveChanges, object originalFormat);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ScreenRefresh();

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOutOld();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOutOld(object background);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOutOld(object background, object append);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOutOld(object background, object append, object range);

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
        void PrintOutOld(object background, object append, object range, object outputFileName);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName);

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
        void PrintOutOld(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839803.aspx </remarks>
        /// <param name="name">string name</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void LookupNameProperties(string name);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192415.aspx </remarks>
        /// <param name="unavailableFont">string unavailableFont</param>
        /// <param name="substituteFont">string substituteFont</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SubstituteFont(string unavailableFont, string substituteFont);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        /// <param name="times">optional object times</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool Repeat(object times);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool Repeat();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845561.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="command">string command</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DDEExecute(Int32 channel, string command);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837295.aspx </remarks>
        /// <param name="app">string app</param>
        /// <param name="topic">string topic</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 DDEInitiate(string app, string topic);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837201.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        /// <param name="data">string data</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DDEPoke(Int32 channel, string item, string data);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837546.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        /// <param name="item">string item</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string DDERequest(Int32 channel, string item);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837904.aspx </remarks>
        /// <param name="channel">Int32 channel</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DDETerminate(Int32 channel);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192053.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DDETerminateAll();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3, object arg4);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845364.aspx </remarks>
        /// <param name="arg1">NetOffice.WordApi.Enums.WdKey arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 BuildKeyCode(NetOffice.WordApi.Enums.WdKey arg1, object arg2, object arg3);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        /// <param name="keyCode2">optional object keyCode2</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string KeyString(Int32 keyCode, object keyCode2);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192163.aspx </remarks>
        /// <param name="keyCode">Int32 keyCode</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string KeyString(Int32 keyCode);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835492.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="destination">string destination</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void OrganizerCopy(string source, string destination, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194744.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void OrganizerDelete(string source, string name, NetOffice.WordApi.Enums.WdOrganizerObject _object);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836140.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="name">string name</param>
        /// <param name="newName">string newName</param>
        /// <param name="_object">NetOffice.WordApi.Enums.WdOrganizerObject object</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void OrganizerRename(string source, string name, string newName, NetOffice.WordApi.Enums.WdOrganizerObject _object);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823266.aspx </remarks>
        /// <param name="tagID">String[] tagID</param>
        /// <param name="value">String[] value</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void AddAddress(String[] tagID, String[] value);

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
        string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice, object updateRecentAddresses);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string GetAddress();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string GetAddress(object name);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string GetAddress(object name, object addressProperties);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836577.aspx </remarks>
        /// <param name="name">optional object name</param>
        /// <param name="addressProperties">optional object addressProperties</param>
        /// <param name="useAutoText">optional object useAutoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string GetAddress(object name, object addressProperties, object useAutoText);

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
        string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog);

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
        string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog);

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
        string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog);

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
        string GetAddress(object name, object addressProperties, object useAutoText, object displaySelectDialog, object selectDialog, object checkNamesDialog, object recentAddressesChoice);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194798.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CheckGrammar(string _string);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CheckSpelling(string word);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CheckSpelling(string word, object customDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822597.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8);

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
        bool CheckSpelling(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822681.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ResetIgnoreAll();

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835170.aspx </remarks>
        /// <param name="word">string word</param>
        /// <param name="customDictionary">optional object customDictionary</param>
        /// <param name="ignoreUppercase">optional object ignoreUppercase</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8);

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
        NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(string word, object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838545.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void GoBack();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841057.aspx </remarks>
        /// <param name="helpType">object helpType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Help(object helpType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194337.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void AutomaticChange();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839095.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ShowMe();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821932.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void HelpTool();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845336.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.Window NewWindow();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194509.aspx </remarks>
        /// <param name="listAllCommands">bool listAllCommands</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ListCommands(bool listAllCommands);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834517.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ShowClipboard();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        /// <param name="tolerance">optional object tolerance</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void OnTime(object when, string name, object tolerance);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820816.aspx </remarks>
        /// <param name="when">object when</param>
        /// <param name="name">string name</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void OnTime(object when, string name);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837154.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void NextLetter();

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
        Int16 MountVolume(string zone, string server, string volume, object user, object userPassword, object volumePassword);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int16 MountVolume(string zone, string server, string volume);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="zone">string zone</param>
        /// <param name="server">string server</param>
        /// <param name="volume">string volume</param>
        /// <param name="user">optional object user</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int16 MountVolume(string zone, string server, string volume, object user);

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
        Int16 MountVolume(string zone, string server, string volume, object user, object userPassword);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844818.aspx </remarks>
        /// <param name="_string">string string</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string CleanString(string _string);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SendFax();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835219.aspx </remarks>
        /// <param name="path">string path</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ChangeFileOpenDirectory(string path);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="macroName">string macroName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void RunOld(string macroName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196922.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void GoForward();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844914.aspx </remarks>
        /// <param name="left">Int32 left</param>
        /// <param name="top">Int32 top</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Move(Int32 left, Int32 top);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197452.aspx </remarks>
        /// <param name="width">Int32 width</param>
        /// <param name="height">Int32 height</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Resize(Int32 width, Int32 height);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197549.aspx </remarks>
        /// <param name="inches">Single inches</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single InchesToPoints(Single inches);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838159.aspx </remarks>
        /// <param name="centimeters">Single centimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single CentimetersToPoints(Single centimeters);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845767.aspx </remarks>
        /// <param name="millimeters">Single millimeters</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single MillimetersToPoints(Single millimeters);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840225.aspx </remarks>
        /// <param name="picas">Single picas</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PicasToPoints(Single picas);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840343.aspx </remarks>
        /// <param name="lines">Single lines</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single LinesToPoints(Single lines);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838268.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToInches(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195052.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToCentimeters(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836929.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToMillimeters(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193434.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToPicas(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822110.aspx </remarks>
        /// <param name="points">Single points</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToLines(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821351.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void Activate();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToPixels(Single points, object fVertical);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840896.aspx </remarks>
        /// <param name="points">Single points</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PointsToPixels(Single points);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        /// <param name="fVertical">optional object fVertical</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PixelsToPoints(Single pixels, object fVertical);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840582.aspx </remarks>
        /// <param name="pixels">Single pixels</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Single PixelsToPoints(Single pixels);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845662.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void KeyboardLatin();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196621.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void KeyboardBidi();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835971.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void ToggleKeyboard();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        /// <param name="langId">optional Int32 LangId = 0</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Keyboard(object langId);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197538.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Keyboard();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193728.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string ProductCode();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840160.aspx </remarks>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.WordApi.DefaultWebOptions DefaultWebOptions();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="range">object range</param>
        /// <param name="cid">object cid</param>
        /// <param name="piCSE">object piCSE</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void DiscussionSupport(object range, object cid, object piCSE);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821531.aspx </remarks>
        /// <param name="name">string name</param>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void SetDefaultTheme(string name, NetOffice.WordApi.Enums.WdDocumentMedium documentType);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834585.aspx </remarks>
        /// <param name="documentType">NetOffice.WordApi.Enums.WdDocumentMedium documentType</param>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        string GetDefaultTheme(NetOffice.WordApi.Enums.WdDocumentMedium documentType);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut();

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object background);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object background, object append);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840681.aspx </remarks>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        void PrintOut(object background, object append, object range);

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
        void PrintOut(object background, object append, object range, object outputFileName);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow);

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
        void PrintOut(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29, object varg30);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        object Run(string macroName);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        object Run(string macroName, object varg1);

        /// <summary>
        /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838935.aspx </remarks>
        /// <param name="macroName">string macroName</param>
        /// <param name="varg1">optional object varg1</param>
        /// <param name="varg2">optional object varg2</param>
        [CustomMethod]
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        object Run(string macroName, object varg1, object varg2);

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
        object Run(string macroName, object varg1, object varg2, object varg3);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28);

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
        object Run(string macroName, object varg1, object varg2, object varg3, object varg4, object varg5, object varg6, object varg7, object varg8, object varg9, object varg10, object varg11, object varg12, object varg13, object varg14, object varg15, object varg16, object varg17, object varg18, object varg19, object varg20, object varg21, object varg22, object varg23, object varg24, object varg25, object varg26, object varg27, object varg28, object varg29);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth, object printZoomPaperHeight);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PrintOut2000();

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PrintOut2000(object background);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PrintOut2000(object background, object append);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="background">optional object background</param>
        /// <param name="append">optional object append</param>
        /// <param name="range">optional object range</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        void PrintOut2000(object background, object append, object range);

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
        void PrintOut2000(object background, object append, object range, object outputFileName);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow);

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
        void PrintOut2000(object background, object append, object range, object outputFileName, object from, object to, object item, object copies, object pages, object pageType, object printToFile, object collate, object fileName, object activePrinterMacGX, object manualDuplexPrint, object printZoomColumn, object printZoomRow, object printZoomPaperWidth);

        /// <summary>
        /// SupportByVersion Word 10, 11, 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 10, 11, 12, 14, 15, 16)]
        bool Dummy2();

        /// <summary>
        /// SupportByVersion Word 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838158.aspx </remarks>
        [SupportByVersion("Word", 11, 12, 14, 15, 16)]
        void PutFocusInMailHeader();

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840673.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 12, 14, 15, 16)]
        void LoadMasterList(string fileName);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor, object ignoreAllComparisonWarnings);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195665.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves);

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
        NetOffice.WordApi.Document CompareDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object revisedAuthor);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor, object formatFrom);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument);

        /// <summary>
        /// SupportByVersion Word 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194674.aspx </remarks>
        /// <param name="originalDocument">NetOffice.WordApi.Document originalDocument</param>
        /// <param name="revisedDocument">NetOffice.WordApi.Document revisedDocument</param>
        /// <param name="destination">optional NetOffice.WordApi.Enums.WdCompareDestination Destination = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 12, 14, 15, 16)]
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor);

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
        NetOffice.WordApi.Document MergeDocuments(NetOffice.WordApi.Document originalDocument, NetOffice.WordApi.Document revisedDocument, object destination, object granularity, object compareFormatting, object compareCaseChanges, object compareWhitespace, object compareTables, object compareHeaders, object compareFootnotes, object compareTextboxes, object compareFields, object compareComments, object compareMoves, object originalAuthor, object revisedAuthor);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="localDocument">NetOffice.WordApi.Document localDocument</param>
        /// <param name="serverDocument">NetOffice.WordApi.Document serverDocument</param>
        /// <param name="baseDocument">NetOffice.WordApi.Document baseDocument</param>
        /// <param name="favorSource">bool favorSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        void ThreeWayMerge(NetOffice.WordApi.Document localDocument, NetOffice.WordApi.Document serverDocument, NetOffice.WordApi.Document baseDocument, bool favorSource);

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        void Dummy4();

        #endregion
    }
}
