using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Shapes : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.Shape>
	{
		#pragma warning disable

		#region Type Information

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
                    _type = typeof(Shapes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Shapes(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Shapes(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Shapes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape Background
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shape>(this, "Background", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape Default
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shape>(this, "Default", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "Item", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddCallout", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddConnector", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddCurve", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddLabel", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddLine", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddPicture", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddPicture", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddPicture", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddPolyline", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddShape", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddTextEffect", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddTextbox", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.OfficeApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.ShapeRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.ShapeRange>(this, "Range", NetOffice.OfficeApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddDiagram", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddCanvas(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddCanvas", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, type, left);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, type, left, top);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddChart(object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, type, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Shape AddTable(Int32 numRows, Int32 numColumns, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddTable", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddSmartArt", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ layout, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddSmartArt", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, layout);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddSmartArt", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, layout, left);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddSmartArt", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, layout, left, top);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddSmartArt", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, layout, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = true</param>
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height, newLayout });
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, style, type);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, style, type, left);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, style, type, left, top);
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Office 15,16
		/// </summary>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 15, 16)]
		public NetOffice.OfficeApi.Shape AddChart2(object style, object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Shape>(this, "AddChart2", NetOffice.OfficeApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height });
		}

        #endregion
       
        #region IEnumerableProvider<NetOffice.OfficeApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Shape>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}