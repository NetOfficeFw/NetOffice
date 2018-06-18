using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOAPPSETTINGS 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIOAPPSETTINGS : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVApplication Application { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisObjectTypes ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool DrawingAids { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthRulerX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthRulerY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGridX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGridY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGuidesX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGuidesY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthPointsX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthPointsY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGeometryX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthGeometryY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthExtensionsX { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 SnapStrengthExtensionsY { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ShowFileSaveWarnings { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ShowFileOpenWarnings { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisDefaultSaveFormats DefaultSaveFormat { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 DrawingPageColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 DrawingBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 DrawingBackgroundColorGradient { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 StencilBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 StencilBackgroundColorGradient { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 StencilTextColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 PrintPreviewBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 FullScreenBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ShowStartupDialog { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ShowSmartTags { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisTextDisplayQualityTypes TextDisplayQuality { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool HigherQualityShapeDisplay { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool SmoothDrawing { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 StencilCharactersPerLine { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 StencilLinesPerMaster { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string UserName { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string UserInitials { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ZoomOnRoll { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 UndoLevels { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 RecentFilesListSize { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool CenterSelectionOnZoom { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ConnectorSplittingEnabled { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions AsianTextUI { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions ComplexTextUI { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions KanaFindAndReplace { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 FreeformDrawingPrecision { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 FreeformDrawingSmoothing { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool DeveloperMode { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool ShowChooseDrawingTypePane { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool ShowShapeSearchPane { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool ApplyThemesOnShapeAdd { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions SATextUI { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions BIDITextUI { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisRegionalUIOptions KashidaTextUI { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool ShowMoreShapeHandlesOnHover { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool EnableAutoConnect { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool ApplyBackgroundToDocument { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool TransitionsEnabled { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool EnableFormulaAutoComplete { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool DeleteConnectorsEnabled { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 RecentTemplatesListSize { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportDataFormat RasterExportDataFormat { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportDataCompression RasterExportDataCompression { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportColorReduction RasterExportColorReduction { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportColorFormat RasterExportColorFormat { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportOperation RasterExportOperation { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportRotation RasterExportRotation { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.Enums.VisRasterExportFlip RasterExportFlip { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 RasterExportBackgroundColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 RasterExportTransparencyColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool RasterExportUseTransparencyColor { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 RasterExportQuality { get; set; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		NetOffice.VisioApi.Enums.VisSVGExportFormat SVGExportFormat { get; set; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		bool EnableLowMemoryMode { get; set; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		bool EnterCommitsText { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		/// <param name="resolutionUnits">optional NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits resolutionUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height, object resolutionUnits);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="resolution">NetOffice.VisioApi.Enums.VisRasterExportResolution resolution</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportResolution(NetOffice.VisioApi.Enums.VisRasterExportResolution resolution, object width, object height);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pResolution">NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pResolutionUnits">NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		void GetRasterExportResolution(out NetOffice.VisioApi.Enums.VisRasterExportResolution pResolution, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportResolutionUnits pResolutionUnits);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		/// <param name="sizeUnits">optional NetOffice.VisioApi.Enums.VisRasterExportSizeUnits sizeUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height, object sizeUnits);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="size">NetOffice.VisioApi.Enums.VisRasterExportSize size</param>
		/// <param name="width">optional Double Width = 0</param>
		/// <param name="height">optional Double Height = 0</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		void SetRasterExportSize(NetOffice.VisioApi.Enums.VisRasterExportSize size, object width, object height);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="pSize">NetOffice.VisioApi.Enums.VisRasterExportSize pSize</param>
		/// <param name="pWidth">Double pWidth</param>
		/// <param name="pHeight">Double pHeight</param>
		/// <param name="pSizeUnits">NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits</param>
		[SupportByVersion("Visio", 14,15,16)]
		void GetRasterExportSize(out NetOffice.VisioApi.Enums.VisRasterExportSize pSize, out Double pWidth, out Double pHeight, out NetOffice.VisioApi.Enums.VisRasterExportSizeUnits pSizeUnits);

		#endregion
	}
}
