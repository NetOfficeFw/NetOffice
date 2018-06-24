using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "MSProject", 11), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("C6984804-2C4D-4874-B69F-C14BF97C0BF1")]
	public interface Shapes : ICOMObject, IEnumerableProvider<NetOffice.MSProjectApi.Shape>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSProject", 11), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		NetOffice.MSProjectApi.Shape get_Value(object index);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Alias for get_Value
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11), Redirect("get_Value")]
		NetOffice.MSProjectApi.Shape Value(object index);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape Background { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape Default { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.MSProjectApi.Shape this[object index] { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddCurve(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddPolyline(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.OfficeApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.ShapeRange Range(object index);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void SelectAll();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = true</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height, object newLayout);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddChart(object style, object type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shape AddTable(Int32 numRows, Int32 numColumns, Single left, Single top, Single width, Single height);

        #endregion

        #region IEnumerable<NetOffice.MSProjectApi.Shape>

        /// <summary>
        /// SupportByVersion MSProject, 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        new IEnumerator<NetOffice.MSProjectApi.Shape> GetEnumerator();

        #endregion
    }
}
