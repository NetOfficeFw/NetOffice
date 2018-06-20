using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746621.aspx </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "PowerPoint", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "Item")]
	[TypeId("91493475-5A91-11CF-8700-00AA0060263B")]
	public interface Shapes : ICOMObject, IEnumerableProvider<NetOffice.PowerPointApi.Shape>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744245.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Application { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744779.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 Creator { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746026.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746106.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743925.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.OfficeApi.Enums.MsoTriState HasTitle { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745054.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape Title { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744297.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Placeholders Placeholders { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.PowerPointApi.Shape this[object index] { get; }

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746541.aspx </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744679.aspx </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744765.aspx </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddCurve(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746034.aspx </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745776.aspx </remarks>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745953.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745953.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745953.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746531.aspx </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPolyline(object safeArrayOfPoints);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744336.aspx </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744722.aspx </remarks>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743980.aspx </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744514.aspx </remarks>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745770.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		void SelectAll();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745017.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange Range(object index);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745017.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange Range();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744198.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTitle();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object link);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745501.aspx </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		/// <param name="width">optional Single Width = 145.25</param>
		/// <param name="height">optional Single Height = 145.25</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddComment(object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddComment();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddComment(object left);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddComment(object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		/// <param name="width">optional Single Width = 145.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddComment(object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744030.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744030.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744030.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744030.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744030.aspx </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745385.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745385.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745385.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745385.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745385.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745532.aspx </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange Paste();

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745319.aspx </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745319.aspx </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745319.aspx </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745319.aspx </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745319.aspx </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object link);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial();

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745158.aspx </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddCanvas(Single left, Single top, Single width, Single height);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart();

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart(object type);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart(object type, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744080.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746195.aspx </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746195.aspx </remarks>
		/// <param name="embedTag">string embedTag</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746195.aspx </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746195.aspx </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746195.aspx </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744968.aspx </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744968.aspx </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744968.aspx </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744968.aspx </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744968.aspx </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = false</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object newLayout);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2();

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width);

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227346.aspx </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width, object height);

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Shape>

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.PowerPointApi.Shape> GetEnumerator();

        #endregion
    }
}
