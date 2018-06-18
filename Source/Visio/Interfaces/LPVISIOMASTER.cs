using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// Interface LPVISIOMASTER 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsInterface)]
	[TypeId("00000000-0000-0000-0000-000000000000")]
	public interface LPVISIOMASTER : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVDocument Document { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Prompt { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 AlignName { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 IconSize { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 IconUpdate { get; set; }

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
		Int16 Stat { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 ObjectType { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShapes Shapes { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Index { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 OneD { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string UniqueID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVLayers Layers { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape PageSheet { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVEventList EventList { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PersistsEvents { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVConnects Connects { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		Int16 ID16 { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVOLEObjects OLEObjects { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 PatternFlags { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 MatchByName { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Hidden { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string BaseID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string NewBaseID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVSelection get_SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Alias for get_SpatialSearch
		/// </summary>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="relation">Int16 relation</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16), Redirect("get_SpatialSearch")]
		NetOffice.VisioApi.IVSelection SpatialSearch(Double x, Double y, Int16 relation, Double tolerance, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		string NameU { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 IndexInStencil { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		stdole.Picture Picture { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		stdole.Picture Icon { get; set; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster EditCopy { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster Original { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		bool IsChanged { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisMasterTypes Type { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool DataGraphicHidden { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool DataGraphicHidesText { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		bool DataGraphicShowBorder { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisGraphicPositionHorizontal DataGraphicHorizontalPosition { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		NetOffice.VisioApi.Enums.VisGraphicPositionVertical DataGraphicVerticalPosition { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVGraphicItems GraphicItems { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void CenterDrawing();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawLine(Double xBegin, Double yBegin, Double xEnd, Double yEnd);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawRectangle(Double x1, Double y1, Double x2, Double y2);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="x1">Double x1</param>
		/// <param name="y1">Double y1</param>
		/// <param name="x2">Double x2</param>
		/// <param name="y2">Double y2</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawOval(Double x1, Double y1, Double x2, Double y2);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="tolerance">Double tolerance</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawSpline(Double[] xyArray, Double tolerance, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawBezier(Double[] xyArray, Int16 degree, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawPolyline(Double[] xyArray, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Import(string fileName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Export(string fileName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape InsertFromFile(string fileName, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="classOrProgID">string classOrProgID</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape InsertObject(string classOrProgID, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVWindow OpenDrawWindow();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVWindow OpenIconWindow();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMaster Open();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Close();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 DropMany(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void GetFormulas(Int16[] sID_SRCStream, out object[] formulaArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void GetResults(Int16[] sID_SRCStream, Int16 flags, object[] unitsNamesOrCodes, out object[] resultArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 SetFormulas(Int16[] sID_SRCStream, object[] formulaArray, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="unitsNamesOrCodes">object[] unitsNamesOrCodes</param>
		/// <param name="resultArray">object[] resultArray</param>
		/// <param name="flags">Int16 flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 SetResults(Int16[] sID_SRCStream, object[] unitsNamesOrCodes, object[] resultArray, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ImportIcon(string fileName);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ExportIconTransparentAsBlack(string fileName, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Layout();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">Int16 flags</param>
		/// <param name="lpr8Left">Double lpr8Left</param>
		/// <param name="lpr8Bottom">Double lpr8Bottom</param>
		/// <param name="lpr8Right">Double lpr8Right</param>
		/// <param name="lpr8Top">Double lpr8Top</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void BoundingBox(Int16 flags, out Double lpr8Left, out Double lpr8Bottom, out Double lpr8Right, out Double lpr8Top);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVMasterShortcut CreateShortcut();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="iDArray">Int16[] iDArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 DropManyU(object[] objectsToInstance, Double[] xyArray, out Int16[] iDArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sID_SRCStream">Int16[] sID_SRCStream</param>
		/// <param name="formulaArray">object[] formulaArray</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void GetFormulasU(Int16[] sID_SRCStream, out object[] formulaArray);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		/// <param name="weights">optional object weights</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots, object weights);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="degree">Int16 degree</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="xyArray">Double[] xyArray</param>
		/// <param name="knots">Double[] knots</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawNURBS(Int16 degree, Int16 flags, Double[] xyArray, Double[] knots);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		/// <param name="transparentRGB">optional object transparentRGB</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ExportIcon(string fileName, Int16 flags, object transparentRGB);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="flags">Int16 flags</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ExportIcon(string fileName, Int16 flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void ResizeToFitContents();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="flags">optional object flags</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Paste(object flags);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Paste();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void PasteSpecial(Int32 format, object link, object displayAsIcon);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void PasteSpecial(Int32 format);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int32 format</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void PasteSpecial(Int32 format, object link);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		/// <param name="data">optional object data</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode, object data);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="selType">NetOffice.VisioApi.Enums.VisSelectionTypes selType</param>
		/// <param name="iterationMode">optional NetOffice.VisioApi.Enums.VisSelectMode IterationMode = 256</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVSelection CreateSelection(NetOffice.VisioApi.Enums.VisSelectionTypes selType, object iterationMode);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">Int16 type</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape AddGuide(Int16 type, Double xPos, Double yPos);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="xControl">Double xControl</param>
		/// <param name="yControl">Double yControl</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawArcByThreePoints(Double xBegin, Double yBegin, Double xEnd, Double yEnd, Double xControl, Double yControl);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xBegin">Double xBegin</param>
		/// <param name="yBegin">Double yBegin</param>
		/// <param name="xEnd">Double xEnd</param>
		/// <param name="yEnd">Double yEnd</param>
		/// <param name="sweepFlag">NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawQuarterArc(Double xBegin, Double yBegin, Double xEnd, Double yEnd, NetOffice.VisioApi.Enums.VisArcSweepFlags sweepFlag);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		/// <param name="endAngle">optional Double EndAngle = 3.1415927410125732</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle, object endAngle);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="xCenter">Double xCenter</param>
		/// <param name="yCenter">Double yCenter</param>
		/// <param name="radius">Double radius</param>
		/// <param name="startAngle">optional Double StartAngle = 0</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.IVShape DrawCircularArc(Double xCenter, Double yCenter, Double radius, object startAngle);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		void DataGraphicDelete();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="flags">Int32 flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void PasteToLocation(Double xPos, Double yPos, Int32 flags);

		#endregion
	}
}
