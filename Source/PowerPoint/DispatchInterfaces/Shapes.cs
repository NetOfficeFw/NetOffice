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
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes"/> </remarks>
	[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Shapes : COMObject, IEnumerableProvider<NetOffice.PowerPointApi.Shape>
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
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Application"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Application
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Creator"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Parent"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Count"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.HasTitle"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasTitle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasTitle");
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Title"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape Title
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shape>(this, "Title", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Placeholders"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Placeholders Placeholders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Placeholders>(this, "Placeholders", NetOffice.PowerPointApi.Placeholders.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PowerPointApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "Item", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddCallout"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddCallout", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddConnector"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddConnector", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddCurve"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddCurve", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddLabel"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddLabel", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddLine"/> </remarks>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddLine", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPicture", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPicture", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">NetOffice.OfficeApi.Enums.MsoTriState linkToFile</param>
		/// <param name="saveWithDocument">NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPicture(string fileName, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, Single left, Single top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPicture", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPolyline"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPolyline", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddShape"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddShape", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTextEffect"/> </remarks>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTextEffect", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTextbox"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTextbox", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.BuildFreeform"/> </remarks>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.PowerPointApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.SelectAll"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Range"/> </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "Range", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Range"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange Range()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "Range", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTitle"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTitle()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTitle", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
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
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object link)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName, displayAsIcon, iconFileName, iconIndex, iconLabel, link });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="className">optional string ClassName = </param>
		/// <param name="fileName">optional string FileName = </param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName, displayAsIcon });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
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
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName, displayAsIcon, iconFileName });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
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
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddOLEObject"/> </remarks>
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
		public NetOffice.PowerPointApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object fileName, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddOLEObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, fileName, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		/// <param name="width">optional Single Width = 145.25</param>
		/// <param name="height">optional Single Height = 145.25</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddComment(object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddComment", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddComment()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddComment", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddComment(object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddComment", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddComment(object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddComment", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">optional Single Left = 1.25</param>
		/// <param name="top">optional Single Top = 1.25</param>
		/// <param name="width">optional Single Width = 145.25</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddComment(object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddComment", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPlaceholder"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPlaceholder", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPlaceholder"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPlaceholder", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPlaceholder"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPlaceholder", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPlaceholder"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPlaceholder", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddPlaceholder"/> </remarks>
		/// <param name="type">NetOffice.PowerPointApi.Enums.PpPlaceholderType type</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddPlaceholder(NetOffice.PowerPointApi.Enums.PpPlaceholderType type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddPlaceholder", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject(string fileName, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.Paste"/> </remarks>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "Paste", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable"/> </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTable", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable"/> </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTable", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, numRows, numColumns);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable"/> </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTable", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, numRows, numColumns, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable"/> </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTable", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, numRows, numColumns, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddTable"/> </remarks>
		/// <param name="numRows">Int32 numRows</param>
		/// <param name="numColumns">Int32 numColumns</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddTable", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		/// <param name="link">optional NetOffice.OfficeApi.Enums.MsoTriState Link = 0</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object link)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, new object[]{ dataType, displayAsIcon, iconFileName, iconIndex, iconLabel, link });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, dataType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, dataType, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, dataType, displayAsIcon, iconFileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, dataType, displayAsIcon, iconFileName, iconIndex);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.PasteSpecial"/> </remarks>
		/// <param name="dataType">optional NetOffice.PowerPointApi.Enums.PpPasteDataType DataType = 0</param>
		/// <param name="displayAsIcon">optional NetOffice.OfficeApi.Enums.MsoTriState DisplayAsIcon = 0</param>
		/// <param name="iconFileName">optional string IconFileName = </param>
		/// <param name="iconIndex">optional Int32 IconIndex = 0</param>
		/// <param name="iconLabel">optional string IconLabel = </param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.ShapeRange PasteSpecial(object dataType, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.ShapeRange>(this, "PasteSpecial", NetOffice.PowerPointApi.ShapeRange.LateBindingApiWrapperType, new object[]{ dataType, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddDiagram", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddCanvas(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddCanvas", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart(object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 12,14,15,16)]
		public NetOffice.PowerPointApi.Shape AddChart(object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, type, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, linkToFile);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObject2"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional NetOffice.OfficeApi.Enums.MsoTriState LinkToFile = 0</param>
		/// <param name="saveWithDocument">optional NetOffice.OfficeApi.Enums.MsoTriState SaveWithDocument = -1</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObject2(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObject2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObjectFromEmbedTag"/> </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObjectFromEmbedTag", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ embedTag, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObjectFromEmbedTag"/> </remarks>
		/// <param name="embedTag">string embedTag</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObjectFromEmbedTag", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, embedTag);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObjectFromEmbedTag"/> </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObjectFromEmbedTag", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, embedTag, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObjectFromEmbedTag"/> </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObjectFromEmbedTag", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, embedTag, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddMediaObjectFromEmbedTag"/> </remarks>
		/// <param name="embedTag">string embedTag</param>
		/// <param name="left">optional Single Left = 0</param>
		/// <param name="top">optional Single Top = 0</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddMediaObjectFromEmbedTag(string embedTag, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddMediaObjectFromEmbedTag", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, embedTag, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddSmartArt", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ layout, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddSmartArt", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, layout);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddSmartArt", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, layout, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddSmartArt", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, layout, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 14,15,16)]
		public NetOffice.PowerPointApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddSmartArt", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, layout, left, top, width);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		/// <param name="newLayout">optional bool NewLayout = false</param>
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height, newLayout });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, style, type);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, style, type, left);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, style, type, left, top);
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width });
		}

		/// <summary>
		/// SupportByVersion PowerPoint 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional Single Left = -1</param>
		/// <param name="top">optional Single Top = -1</param>
		/// <param name="width">optional Single Width = -1</param>
		/// <param name="height">optional Single Height = -1</param>
		[CustomMethod]
		[SupportByVersion("PowerPoint", 15, 16)]
		public NetOffice.PowerPointApi.Shape AddChart2(object style, object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PowerPointApi.Shape>(this, "AddChart2", NetOffice.PowerPointApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height });
		}

        #endregion
       
        #region IEnumerableProvider<NetOffice.PowerPointApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.PowerPointApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.PowerPointApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PowerPointApi.Shape>

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.PowerPointApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PowerPointApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("PowerPoint", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}