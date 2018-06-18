using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi
{
	/// <summary>
	/// DispatchInterface IVPage 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("000D0709-0000-0000-C000-000000000046")]
    [CoClassSource(typeof(NetOffice.VisioApi.Page))]
    public interface IVPage : ICOMObject
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int16 Background { get; set; }

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
		Int16 Index { get; set; }

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
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.VisioApi.IVPage BackPageAsObj { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string BackPageFromName { get; set; }

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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		object BackPage { get; set; }

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
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ID { get; }

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
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16), NativeResult]
		stdole.Picture Picture { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 PrintTileCount { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		NetOffice.VisioApi.Enums.VisPageTypes Type { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		Int32 ReviewerID { get; }

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVPage OriginalPage { get; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		object ThemeColors { get; set; }

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 12,14,15,16)]
		object ThemeEffects { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool LayoutRoutePassive { get; set; }

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		bool AutoSize { get; set; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVComments Comments { get; }

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// Get
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVComments ShapeComments { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void old_Paste();

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="format">Int16 format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void old_PasteSpecial(Int16 format);

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
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape Drop(object objectToDrop, Double xPos, Double yPos);

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
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Print();

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
		/// <param name="fRenumberPages">Int16 fRenumberPages</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void Delete(Int16 fRenumberPages);

		/// <summary>
		/// SupportByVersion Visio 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void CenterDrawing();

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
		/// <param name="nTile">Int32 nTile</param>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		void PrintTile(Int32 nTile);

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
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetShapesLinkedToData(Int32 dataRecordsetID, out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void GetShapesLinkedToDataRow(Int32 dataRecordsetID, Int32 dataRowID, out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		/// <param name="applyDataGraphicAfterLink">optional bool ApplyDataGraphicAfterLink = true</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs, object applyDataGraphicAfterLink);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[CustomMethod]
		[SupportByVersion("Visio", 12,14,15,16)]
		void LinkShapesToDataRows(Int32 dataRecordsetID, Int32[] dataRowIDs, Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		/// <param name="uniqueIDArgs">NetOffice.VisioApi.Enums.VisUniqueIDArgs uniqueIDArgs</param>
		/// <param name="gUIDs">String[] gUIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void ShapeIDsToUniqueIDs(Int32[] shapeIDs, NetOffice.VisioApi.Enums.VisUniqueIDArgs uniqueIDArgs, out String[] gUIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="gUIDs">String[] gUIDs</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		void UniqueIDsToShapeIDs(String[] gUIDs, out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="x">Double x</param>
		/// <param name="y">Double y</param>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowID">Int32 dataRowID</param>
		/// <param name="applyDataGraphicAfterLink">bool applyDataGraphicAfterLink</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropLinked(object objectToDrop, Double x, Double y, Int32 dataRecordsetID, Int32 dataRowID, bool applyDataGraphicAfterLink);

		/// <summary>
		/// SupportByVersion Visio 12, 14, 15, 16
		/// </summary>
		/// <param name="objectsToInstance">object[] objectsToInstance</param>
		/// <param name="xYs">Double[] xYs</param>
		/// <param name="dataRecordsetID">Int32 dataRecordsetID</param>
		/// <param name="dataRowIDs">Int32[] dataRowIDs</param>
		/// <param name="applyDataGraphicAfterLink">bool applyDataGraphicAfterLink</param>
		/// <param name="shapeIDs">Int32[] shapeIDs</param>
		[SupportByVersion("Visio", 12,14,15,16)]
		Int32 DropManyLinkedU(object[] objectsToInstance, Double[] xYs, Int32 dataRecordsetID, Int32[] dataRowIDs, bool applyDataGraphicAfterLink, out Int32[] shapeIDs);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir, object connector);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		/// <param name="placementDir">NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir</param>
		[CustomMethod]
		[BaseResult]
		[SupportByVersion("Visio", 14,15,16)]
		NetOffice.VisioApi.IVShape DropConnected(object objectToDrop, NetOffice.VisioApi.IVShape targetShape, NetOffice.VisioApi.Enums.VisAutoConnectDir placementDir);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] fromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] toShapeIDs</param>
		/// <param name="placementDirs">Int32[] placementDirs</param>
		/// <param name="connector">optional object Connector = null (Nothing in visual basic)</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs, object connector);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="fromShapeIDs">Int32[] fromShapeIDs</param>
		/// <param name="toShapeIDs">Int32[] toShapeIDs</param>
		/// <param name="placementDirs">Int32[] placementDirs</param>
		[CustomMethod]
		[SupportByVersion("Visio", 14,15,16)]
		Int32 AutoConnectMany(Int32[] fromShapeIDs, Int32[] toShapeIDs, Int32[] placementDirs);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShapes">object targetShapes</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropContainer(object objectToDrop, object targetShapes);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="alignOrSpace">NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace</param>
		/// <param name="alignHorizontal">NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal</param>
		/// <param name="alignVertical">NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical</param>
		/// <param name="spaceHorizontal">Double spaceHorizontal</param>
		/// <param name="spaceVertical">Double spaceVertical</param>
		/// <param name="unitsNameOrCode">NetOffice.VisioApi.Enums.VisUnitCodes unitsNameOrCode</param>
		[SupportByVersion("Visio", 14,15,16)]
		void LayoutIncremental(NetOffice.VisioApi.Enums.VisLayoutIncrementalType alignOrSpace, NetOffice.VisioApi.Enums.VisLayoutHorzAlignType alignHorizontal, NetOffice.VisioApi.Enums.VisLayoutVertAlignType alignVertical, Double spaceHorizontal, Double spaceVertical, NetOffice.VisioApi.Enums.VisUnitCodes unitsNameOrCode);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="direction">NetOffice.VisioApi.Enums.VisLayoutDirection direction</param>
		[SupportByVersion("Visio", 14,15,16)]
		void LayoutChangeDirection(NetOffice.VisioApi.Enums.VisLayoutDirection direction);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void AvoidPageBreaks();

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="connectorToSplit">NetOffice.VisioApi.IVShape connectorToSplit</param>
		/// <param name="shape">NetOffice.VisioApi.IVShape shape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape SplitConnector(NetOffice.VisioApi.IVShape connectorToSplit, NetOffice.VisioApi.IVShape shape);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetShape">NetOffice.VisioApi.IVShape targetShape</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropCallout(object objectToDrop, NetOffice.VisioApi.IVShape targetShape);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="xPos">Double xPos</param>
		/// <param name="yPos">Double yPos</param>
		/// <param name="flags">Int32 flags</param>
		[SupportByVersion("Visio", 14,15,16)]
		void PasteToLocation(Double xPos, Double yPos, Int32 flags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetContainers(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="nestedOptions">NetOffice.VisioApi.Enums.VisContainerNested nestedOptions</param>
		[SupportByVersion("Visio", 14,15,16)]
		Int32[] GetCallouts(NetOffice.VisioApi.Enums.VisContainerNested nestedOptions);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="outerList">object outerList</param>
		/// <param name="innerContainer">object innerContainer</param>
		/// <param name="populateFlags">NetOffice.VisioApi.Enums.VisLegendFlags populateFlags</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropLegend(object outerList, object innerContainer, NetOffice.VisioApi.Enums.VisLegendFlags populateFlags);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		/// <param name="objectToDrop">object objectToDrop</param>
		/// <param name="targetList">NetOffice.VisioApi.IVShape targetList</param>
		/// <param name="lPosition">Int32 lPosition</param>
		[SupportByVersion("Visio", 14,15,16)]
		[BaseResult]
		NetOffice.VisioApi.IVShape DropIntoList(object objectToDrop, NetOffice.VisioApi.IVShape targetList, Int32 lPosition);

		/// <summary>
		/// SupportByVersion Visio 14, 15, 16
		/// </summary>
		[SupportByVersion("Visio", 14,15,16)]
		void AutoSizeDrawing();

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		[BaseResult]
		NetOffice.VisioApi.IVPage Duplicate();

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="eThemeType">NetOffice.VisioApi.Enums.VisThemeTypes eThemeType</param>
		[SupportByVersion("Visio", 15, 16)]
		object GetTheme(NetOffice.VisioApi.Enums.VisThemeTypes eThemeType);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		/// <param name="varFontScheme">optional object varFontScheme</param>
		[SupportByVersion("Visio", 15, 16)]
		void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme, object varFontScheme);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void SetTheme(object varThemeIndex);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void SetTheme(object varThemeIndex, object varColorScheme);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="varThemeIndex">object varThemeIndex</param>
		/// <param name="varColorScheme">optional object varColorScheme</param>
		/// <param name="varEffectScheme">optional object varEffectScheme</param>
		/// <param name="varConnectorScheme">optional object varConnectorScheme</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void SetTheme(object varThemeIndex, object varColorScheme, object varEffectScheme, object varConnectorScheme);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		/// <param name="pEmbellishment">optional Int16 pEmbellishment = 0</param>
		[SupportByVersion("Visio", 15, 16)]
		void GetThemeVariant(out Int16 pVariantColor, out Int16 pVariantStyle, object pEmbellishment);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="pVariantColor">Int16 pVariantColor</param>
		/// <param name="pVariantStyle">Int16 pVariantStyle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void GetThemeVariant(out Int16 pVariantColor, out Int16 pVariantStyle);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		/// <param name="embellishment">optional Int16 embellishment = -1</param>
		[SupportByVersion("Visio", 15, 16)]
		void SetThemeVariant(Int16 variantColor, Int16 variantStyle, object embellishment);

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		/// <param name="variantColor">Int16 variantColor</param>
		/// <param name="variantStyle">Int16 variantStyle</param>
		[CustomMethod]
		[SupportByVersion("Visio", 15, 16)]
		void SetThemeVariant(Int16 variantColor, Int16 variantStyle);

		#endregion
	}
}
