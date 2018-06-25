using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Publisher", 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("00021235-0000-0000-C000-000000000046")]
	public interface Shapes : ICOMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Shape>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Enums.pbCanvasArrangementType CanvasArrangementType { get; set; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		Int32 CanvasesCount { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PublisherApi.Shape this[object index] { get; }

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, object beginX, object beginY, object endX, object endY);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddCurve(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddLabel(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddLine(object beginX, object beginY, object endX, object endY);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename, object link);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="filename">optional string Filename = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddPolyline(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		/// <param name="direction">optional NetOffice.PublisherApi.Enums.PbTableDirectionType Direction = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize, object direction);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="fixedSize">optional bool FixedSize = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddTextbox(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">object fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		/// <param name="launchPropertiesWindow">optional bool LaunchPropertiesWindow = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height, object launchPropertiesWindow);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.PublisherApi.Enums.PbWebControlType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange Paste();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange Range(object index);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange Range();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		void SelectAll();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		/// <param name="design">optional Int32 Design = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height, object design);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object width</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top, object width);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddCatalogMergeArea();

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="canvasId">Int32 canvasId</param>
		/// <param name="catalogMergeFieldType">NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType</param>
		/// <param name="dbCol">Int32 dbCol</param>
		[SupportByVersion("Publisher", 14,15,16)]
		void AddCatalogMergeFieldToCanvas(Int32 canvasId, NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType, Int32 dbCol);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="presetWordArt">NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">object fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddWordArt(NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top);

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bBlockIn">NetOffice.PublisherApi.BuildingBlock bBlockIn</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		NetOffice.PublisherApi.Shape AddBuildingBlock(NetOffice.PublisherApi.BuildingBlock bBlockIn, object left, object top);

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        new IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator();

        #endregion
    }
}
