using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface _Presentation 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("9149349D-5A91-11CF-8700-00AA0060263B")]
    [CoClassSource(typeof(NetOffice.PowerPointApi.Presentation))]
    public interface _Presentation : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745080.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743905.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745484.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.PowerPointApi._Master SlideMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746378.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.PowerPointApi._Master TitleMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745657.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HasTitleMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744870.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string TemplateName { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743938.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.PowerPointApi._Master NotesMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746405.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.PowerPointApi._Master HandoutMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746142.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Slides Slides { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745413.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.PageSetup PageSetup { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744763.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ColorSchemes ColorSchemes { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746216.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ExtraColors ExtraColors { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745621.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.SlideShowSettings SlideShowSettings { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744620.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Fonts Fonts { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746292.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DocumentWindows Windows { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744602.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Tags Tags { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744397.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape DefaultShape { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746376.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object BuiltInDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744661.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object CustomDocumentProperties { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745299.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.VBIDEApi.VBProject VBProject { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746329.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState ReadOnly { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746313.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string FullName { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745890.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745125.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string Path { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744884.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState Saved { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744109.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpDirection LayoutDirection { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744540.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.PrintOptions PrintOptions { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746277.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Container { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745979.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState DisplayComments { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746803.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpFarEastLineBreakLevel FarEastLineBreakLevel { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746404.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string NoLineBreakBefore { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746110.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		string NoLineBreakAfter { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745765.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.SlideShowWindow SlideShowWindow { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746489.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoFarEastLineBreakLanguageID FarEastLineBreakLanguage { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745465.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoLanguageID DefaultLanguageID { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746786.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.CommandBars CommandBars { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.PublishObjects PublishObjects { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.WebOptions WebOptions { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.HTMLProject HTMLProject { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746516.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState EnvelopeVisible { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746671.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState VBASigned { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746323.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState SnapToGrid { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744975.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		Single GridDistance { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744959.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Designs Designs { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745705.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.SignatureSet Signatures { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746134.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState RemovePersonalInformation { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Enums.PpRevisionInfo HasRevisionInfo { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743904.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		string PasswordEncryptionProvider { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745251.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		string PasswordEncryptionAlgorithm { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744792.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		Int32 PasswordEncryptionKeyLength { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743937.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		bool PasswordEncryptionFileProperties { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745703.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		string Password { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744704.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		string WritePassword { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744658.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		NetOffice.OfficeApi.Permission Permission { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745343.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		NetOffice.OfficeApi.SharedWorkspace SharedWorkspace { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745948.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		NetOffice.OfficeApi.Sync Sync { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744818.aspx </remarks>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		NetOffice.OfficeApi.DocumentLibraryVersions DocumentLibraryVersions { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745118.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.MetaProperties ContentTypeProperties { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		Int32 SectionCount { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		bool HasSections { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745108.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.ServerPolicy ServerPolicy { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744654.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.DocumentInspectors DocumentInspectors { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746792.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		bool HasVBProject { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745253.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.CustomXMLParts CustomXMLParts { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743879.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		bool Final { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745858.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.CustomerData CustomerData { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746518.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Research Research { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745747.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		string EncryptionProvider { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744806.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.SectionProperties SectionProperties { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745326.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Coauthoring Coauthoring { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746423.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool InMergeMode { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744901.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Broadcast Broadcast { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745775.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool HasNotesMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744680.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		bool HasHandoutMaster { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743993.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Enums.PpMediaTaskStatus CreateVideoStatus { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229098.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		bool ChartDataPointTrack { get; set; }

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229702.aspx </remarks>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Guides Guides { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746001.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[BaseResult]
		NetOffice.PowerPointApi._Master AddTitleMaster();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743876.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void ApplyTemplate(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744342.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.DocumentWindow NewWindow();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		/// <param name="headerInfo">optional string HeaderInfo = </param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744691.aspx </remarks>
		/// <param name="address">string address</param>
		/// <param name="subAddress">optional string SubAddress = </param>
		/// <param name="newWindow">optional bool NewWindow = false</param>
		/// <param name="addHistory">optional bool AddHistory = true</param>
		/// <param name="extraInfo">optional string ExtraInfo = </param>
		/// <param name="method">optional NetOffice.OfficeApi.Enums.MsoExtraInfoMethod Method = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void FollowHyperlink(string address, object subAddress, object newWindow, object addHistory, object extraInfo, object method);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744969.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void AddToFavorites();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Unused();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		/// <param name="collate">optional NetOffice.OfficeApi.Enums.MsoTriState Collate = -99</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut(object from, object to, object printToFile, object copies, object collate);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut(object from);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut(object from, object to);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut(object from, object to, object printToFile);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744928.aspx </remarks>
		/// <param name="from">optional Int32 From = -1</param>
		/// <param name="to">optional Int32 To = -1</param>
		/// <param name="printToFile">optional string PrintToFile = </param>
		/// <param name="copies">optional Int32 Copies = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void PrintOut(object from, object to, object printToFile, object copies);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745194.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Save();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveAs(string fileName, object fileFormat, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveAs(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746389.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveAs(string fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		/// <param name="embedTrueTypeFonts">optional NetOffice.OfficeApi.Enums.MsoTriState EmbedTrueTypeFonts = -2</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveCopyAs(string fileName, object fileFormat, object embedTrueTypeFonts);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveCopyAs(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744735.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="fileFormat">optional NetOffice.PowerPointApi.Enums.PpSaveAsFileType FileFormat = 11</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SaveCopyAs(string fileName, object fileFormat);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		/// <param name="scaleHeight">optional Int32 ScaleHeight = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Export(string path, string filterName, object scaleWidth, object scaleHeight);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Export(string path, string filterName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746498.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="filterName">string filterName</param>
		/// <param name="scaleWidth">optional Int32 ScaleWidth = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Export(string path, string filterName, object scaleWidth);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743857.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="text">string text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SetUndoText(string text);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744167.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void UpdateLinks();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void WebPagePreview();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="cp">NetOffice.OfficeApi.Enums.MsoEncoding cp</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void ReloadAs(NetOffice.OfficeApi.Enums.MsoEncoding cp);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="isDesignTemplate">NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void MakeIntoTemplate(NetOffice.OfficeApi.Enums.MsoTriState isDesignTemplate);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="s">string s</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void sblt(string s);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228409.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void Merge(string path);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges, object comments, object makePublic);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void CheckIn();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745069.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void CheckIn(object saveChanges, object comments);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744274.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		bool CanCheckIn();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		/// <param name="includeAttachment">optional object includeAttachment</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject, object showMessage, object includeAttachment);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SendForReview();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SendForReview(object recipients);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SendForReview(object recipients, object subject, object showMessage);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="showMessage">optional bool ShowMessage = true</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void ReplyWithChanges(object showMessage);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void ReplyWithChanges();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746226.aspx </remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void EndReview();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">optional string FileName = </param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void AddBaseline(object fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void AddBaseline();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void RemoveBaseline();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743880.aspx </remarks>
		/// <param name="passwordEncryptionProvider">string passwordEncryptionProvider</param>
		/// <param name="passwordEncryptionAlgorithm">string passwordEncryptionAlgorithm</param>
		/// <param name="passwordEncryptionKeyLength">Int32 passwordEncryptionKeyLength</param>
		/// <param name="passwordEncryptionFileProperties">bool passwordEncryptionFileProperties</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		void SetPasswordEncryptionOptions(string passwordEncryptionProvider, string passwordEncryptionAlgorithm, Int32 passwordEncryptionKeyLength, bool passwordEncryptionFileProperties);

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		/// <param name="showMessage">optional bool ShowMessage = false</param>
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients, object subject, object showMessage);

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		void SendFaxOverInternet();

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients);

		/// <summary>
		/// SupportByVersion PowerPoint 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744724.aspx </remarks>
		/// <param name="recipients">optional string Recipients = </param>
		/// <param name="subject">optional string Subject = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 11,12,14,15,16)]
		void SendFaxOverInternet(object recipients, object subject);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		/// <param name="afterSlide">bool afterSlide</param>
		/// <param name="sectionTitle">string sectionTitle</param>
		/// <param name="newSectionIndex">Int32 newSectionIndex</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void NewSectionAfter(Int32 index, bool afterSlide, string sectionTitle, out Int32 newSectionIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void DeleteSection(Int32 index);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void DisableSections();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		string sectionTitle(Int32 index);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744345.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void RemoveDocumentInformation(NetOffice.PowerPointApi.Enums.PpRemoveDocInfoType type);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		/// <param name="versionType">optional object versionType</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments, object makePublic, object versionType);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void CheckInWithVersion();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746800.aspx </remarks>
		/// <param name="saveChanges">optional bool SaveChanges = true</param>
		/// <param name="comments">optional object comments</param>
		/// <param name="makePublic">optional object makePublic</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void CheckInWithVersion(object saveChanges, object comments, object makePublic);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object externalExporter);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746080.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ExportAsFixedFormat(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746373.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.WorkflowTasks GetWorkflowTasks();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746712.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.OfficeApi.WorkflowTemplates GetWorkflowTemplates();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744831.aspx </remarks>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void LockServerFile();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746528.aspx </remarks>
		/// <param name="themeName">string themeName</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void ApplyTheme(string themeName);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		/// <param name="useSlideOrder">optional bool UseSlideOrder = false</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void PublishSlides(string slideLibraryUrl, object overwrite, object useSlideOrder);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void PublishSlides(string slideLibraryUrl);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744375.aspx </remarks>
		/// <param name="slideLibraryUrl">string slideLibraryUrl</param>
		/// <param name="overwrite">optional bool Overwrite = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void PublishSlides(string slideLibraryUrl, object overwrite);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		void Convert();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744063.aspx </remarks>
		/// <param name="withPresentation">string withPresentation</param>
		/// <param name="baselinePresentation">string baselinePresentation</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void MergeWithBaseline(string withPresentation, string baselinePresentation);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745418.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void AcceptAll();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745993.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void RejectAll();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744528.aspx </remarks>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void EnsureAllMediaUpgraded();

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743830.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void Convert2(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		/// <param name="quality">optional Int32 Quality = 85</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond, object quality);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName, object useTimingsAndNarrations);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746354.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="useTimingsAndNarrations">optional bool UseTimingsAndNarrations = true</param>
		/// <param name="defaultSlideDuration">optional Int32 DefaultSlideDuration = 5</param>
		/// <param name="vertResolution">optional Int32 VertResolution = 720</param>
		/// <param name="framesPerSecond">optional Int32 FramesPerSecond = 30</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		void CreateVideo(string fileName, object useTimingsAndNarrations, object defaultSlideDuration, object vertResolution, object framesPerSecond);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228125.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="variant">string variant</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		void ApplyTemplate2(string fileName, string variant);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="includeMarkup">optional bool IncludeMarkup = false</param>
		/// <param name="externalExporter">optional object externalExporter</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup, object externalExporter);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229507.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="fixedFormatType">NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType</param>
		/// <param name="intent">optional NetOffice.PowerPointApi.Enums.PpFixedFormatIntent Intent = 1</param>
		/// <param name="frameSlides">optional NetOffice.OfficeApi.Enums.MsoTriState FrameSlides = 0</param>
		/// <param name="handoutOrder">optional NetOffice.PowerPointApi.Enums.PpPrintHandoutOrder HandoutOrder = 1</param>
		/// <param name="outputType">optional NetOffice.PowerPointApi.Enums.PpPrintOutputType OutputType = 1</param>
		/// <param name="printHiddenSlides">optional NetOffice.OfficeApi.Enums.MsoTriState PrintHiddenSlides = 0</param>
		/// <param name="printRange">optional NetOffice.PowerPointApi.PrintRange PrintRange = 0</param>
		/// <param name="rangeType">optional NetOffice.PowerPointApi.Enums.PpPrintRangeType RangeType = 1</param>
		/// <param name="slideShowName">optional string SlideShowName = </param>
		/// <param name="includeDocProperties">optional bool IncludeDocProperties = false</param>
		/// <param name="keepIRMSettings">optional bool KeepIRMSettings = true</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="includeMarkup">optional bool IncludeMarkup = false</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		void ExportAsFixedFormat2(string path, NetOffice.PowerPointApi.Enums.PpFixedFormatType fixedFormatType, object intent, object frameSlides, object handoutOrder, object outputType, object printHiddenSlides, object printRange, object rangeType, object slideShowName, object includeDocProperties, object keepIRMSettings, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object includeMarkup);

		#endregion
	}
}
