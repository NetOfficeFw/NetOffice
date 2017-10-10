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
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Shapes : COMObject, NetOffice.CollectionsGeneric.IEnumerableProvider<NetOffice.PublisherApi.Shape>
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Enums.pbCanvasArrangementType CanvasArrangementType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.pbCanvasArrangementType>(this, "CanvasArrangementType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CanvasArrangementType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 CanvasesCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CanvasesCount");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.PublisherApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "Item", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCallout", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, object beginX, object beginY, object endX, object endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddConnector", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCurve", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddLabel(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddLabel", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="beginX">object beginX</param>
		/// <param name="beginY">object beginY</param>
		/// <param name="endX">object endX</param>
		/// <param name="endY">object endY</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddLine(object beginX, object beginY, object endX, object endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddLine", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

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
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename, object link)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, filename, link });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

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
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className });
		}

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
		public NetOffice.PublisherApi.Shape AddOLEObject(object left, object top, object width, object height, object className, object filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddOLEObject", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, className, filename });
		}

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
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top, width, height });
		}

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
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top });
		}

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
		public NetOffice.PublisherApi.Shape AddPicture(string filename, NetOffice.OfficeApi.Enums.MsoTriState linkToFile, NetOffice.OfficeApi.Enums.MsoTriState saveWithDocument, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPicture", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ filename, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddPolyline", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoAutoShapeType type</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddShape(NetOffice.OfficeApi.Enums.MsoAutoShapeType type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddShape", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

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
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize, object direction)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height, fixedSize, direction });
		}

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
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height });
		}

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
		public NetOffice.PublisherApi.Shape AddTable(Int32 numRows, Int32 numColumns, object left, object top, object width, object height, object fixedSize)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTable", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ numRows, numColumns, left, top, width, height, fixedSize });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="orientation">NetOffice.PublisherApi.Enums.PbTextOrientation orientation</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">object width</param>
		/// <param name="height">object height</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddTextbox(NetOffice.PublisherApi.Enums.PbTextOrientation orientation, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTextbox", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

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
		public NetOffice.PublisherApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddTextEffect", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

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
		public NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height, object launchPropertiesWindow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebControl", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height, launchPropertiesWindow });
		}

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
		public NetOffice.PublisherApi.Shape AddWebControl(NetOffice.PublisherApi.Enums.PbWebControlType type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebControl", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">object x1</param>
		/// <param name="y1">object y1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, object x1, object y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.PublisherApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Paste()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Paste", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange Range()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "Range", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

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
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height, object design)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ wizard, left, top, width, height, design });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, wizard, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizard">NetOffice.PublisherApi.Enums.PbWizardGroup wizard</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, wizard, left, top, width);
		}

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
		public NetOffice.PublisherApi.Shape AddGroupWizard(NetOffice.PublisherApi.Enums.PbWizardGroup wizard, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddGroupWizard", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ wizard, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		/// <param name="instance">optional Int32 Instance = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag, object instance)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, wizardTag, instance);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="wizardTag">NetOffice.PublisherApi.Enums.PbWizardTag wizardTag</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ShapeRange FindShapeByWizardTag(NetOffice.PublisherApi.Enums.PbWizardTag wizardTag)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.ShapeRange>(this, "FindShapeByWizardTag", NetOffice.PublisherApi.ShapeRange.LateBindingApiWrapperType, wizardTag);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object width</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebNavigationBar", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, name, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddWebNavigationBar(string name, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWebNavigationBar", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, name, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddCatalogMergeArea()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddCatalogMergeArea", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		/// <param name="height">optional object Height = -1</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		/// <param name="width">optional object Width = -1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddEmptyPictureFrame(object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddEmptyPictureFrame", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="canvasId">Int32 canvasId</param>
		/// <param name="catalogMergeFieldType">NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType</param>
		/// <param name="dbCol">Int32 dbCol</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public void AddCatalogMergeFieldToCanvas(Int32 canvasId, NetOffice.PublisherApi.Enums.pbCatalogMergeFieldType catalogMergeFieldType, Int32 dbCol)
		{
			 Factory.ExecuteMethod(this, "AddCatalogMergeFieldToCanvas", canvasId, catalogMergeFieldType, dbCol);
		}

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
		public NetOffice.PublisherApi.Shape AddWordArt(NetOffice.PublisherApi.Enums.pbPresetWordArt presetWordArt, string text, string fontName, object fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddWordArt", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, new object[]{ presetWordArt, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="bBlockIn">NetOffice.PublisherApi.BuildingBlock bBlockIn</param>
		/// <param name="left">object left</param>
		/// <param name="top">object top</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Shape AddBuildingBlock(NetOffice.PublisherApi.BuildingBlock bBlockIn, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.PublisherApi.Shape>(this, "AddBuildingBlock", NetOffice.PublisherApi.Shape.LateBindingApiWrapperType, bBlockIn, left, top);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.PublisherApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.PublisherApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.PublisherApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.PublisherApi.Shape>

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14, 15, 16)]
        public IEnumerator<NetOffice.PublisherApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.PublisherApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Publisher, 14,15,16
        /// </summary>
        [SupportByVersion("Publisher", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}