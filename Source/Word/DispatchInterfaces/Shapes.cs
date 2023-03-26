using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Shapes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Shapes : COMObject, IEnumerableProvider<NetOffice.WordApi.Shape>
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Shape this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "Item", NetOffice.WordApi.Shape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCallout"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCallout", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCallout"/> </remarks>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoCalloutType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCallout(NetOffice.OfficeApi.Enums.MsoCalloutType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCallout", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoConnectorType type</param>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddConnector(NetOffice.OfficeApi.Enums.MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddConnector", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, beginX, beginY, endX, endY });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCurve"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCurve(object safeArrayOfPoints, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCurve", NetOffice.WordApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints, anchor);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCurve"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCurve(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCurve", NetOffice.WordApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddLabel"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddLabel", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddLabel"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddLabel(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddLabel", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddLine"/> </remarks>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddLine", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ beginX, beginY, endX, endY, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddLine"/> </remarks>
		/// <param name="beginX">Single beginX</param>
		/// <param name="beginY">Single beginY</param>
		/// <param name="endX">Single endX</param>
		/// <param name="endY">Single endY</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddLine(Single beginX, Single beginY, Single endX, Single endY)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddLine", NetOffice.WordApi.Shape.LateBindingApiWrapperType, beginX, beginY, endX, endY);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, fileName, linkToFile);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument, left);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPicture(string fileName, object linkToFile, object saveWithDocument, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPicture", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ fileName, linkToFile, saveWithDocument, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPolyline"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPolyline(object safeArrayOfPoints, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPolyline", NetOffice.WordApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints, anchor);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddPolyline"/> </remarks>
		/// <param name="safeArrayOfPoints">object safeArrayOfPoints</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddPolyline(object safeArrayOfPoints)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddPolyline", NetOffice.WordApi.Shape.LateBindingApiWrapperType, safeArrayOfPoints);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddShape"/> </remarks>
		/// <param name="type">Int32 type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddShape(Int32 type, Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddShape", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddShape"/> </remarks>
		/// <param name="type">Int32 type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddShape(Int32 type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddShape", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddTextEffect"/> </remarks>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddTextEffect", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddTextEffect"/> </remarks>
		/// <param name="presetTextEffect">NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect</param>
		/// <param name="text">string text</param>
		/// <param name="fontName">string fontName</param>
		/// <param name="fontSize">Single fontSize</param>
		/// <param name="fontBold">NetOffice.OfficeApi.Enums.MsoTriState fontBold</param>
		/// <param name="fontItalic">NetOffice.OfficeApi.Enums.MsoTriState fontItalic</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddTextEffect(NetOffice.OfficeApi.Enums.MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, NetOffice.OfficeApi.Enums.MsoTriState fontBold, NetOffice.OfficeApi.Enums.MsoTriState fontItalic, Single left, Single top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddTextEffect", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddTextbox"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddTextbox", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddTextbox"/> </remarks>
		/// <param name="orientation">NetOffice.OfficeApi.Enums.MsoTextOrientation orientation</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddTextbox(NetOffice.OfficeApi.Enums.MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddTextbox", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ orientation, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.BuildFreeform"/> </remarks>
		/// <param name="editingType">NetOffice.OfficeApi.Enums.MsoEditingType editingType</param>
		/// <param name="x1">Single x1</param>
		/// <param name="y1">Single y1</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FreeformBuilder BuildFreeform(NetOffice.OfficeApi.Enums.MsoEditingType editingType, Single x1, Single y1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.FreeformBuilder>(this, "BuildFreeform", NetOffice.WordApi.FreeformBuilder.LateBindingApiWrapperType, editingType, x1, y1);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.Range"/> </remarks>
		/// <param name="index">object index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ShapeRange Range(object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.ShapeRange>(this, "Range", NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType, index);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.SelectAll"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectAll()
		{
			 Factory.ExecuteMethod(this, "SelectAll");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, fileName, linkToFile);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, fileName, linkToFile, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, left });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEObject", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, left);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, left, top);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, classType, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddOLEControl(object classType, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddOLEControl", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ classType, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddDiagram", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.OfficeApi.Enums.MsoDiagramType type</param>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddDiagram(NetOffice.OfficeApi.Enums.MsoDiagramType type, Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddDiagram", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCanvas"/> </remarks>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCanvas(Single left, Single top, Single width, Single height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCanvas", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddCanvas"/> </remarks>
		/// <param name="left">Single left</param>
		/// <param name="top">Single top</param>
		/// <param name="width">Single width</param>
		/// <param name="height">Single height</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Shape AddCanvas(Single left, Single top, Single width, Single height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddCanvas", NetOffice.WordApi.Shape.LateBindingApiWrapperType, left, top, width, height);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, type, left);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, type, left, top);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, type, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Shape AddChart(object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ layout, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, layout);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, layout, left);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, layout, left, top);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, layout, left, top, width);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Shapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Shape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddSmartArt", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ layout, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, left, top, width, height, anchor });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, embedCode, videoWidth, videoHeight);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, embedCode, videoWidth, videoHeight, posterFrameImage);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, left });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, left, top });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddWebVideo", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		/// <param name="newLayout">optional object newLayout</param>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object anchor, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height, anchor, newLayout });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, style, type);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, style, type, left);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left, object top)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, style, type, left, top);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left, object top, object width)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left, object top, object width, object height)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.shapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="left">optional object left</param>
		/// <param name="top">optional object top</param>
		/// <param name="width">optional object width</param>
		/// <param name="height">optional object height</param>
		/// <param name="anchor">optional object anchor</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Shape AddChart2(object style, object type, object left, object top, object width, object height, object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Shape>(this, "AddChart2", NetOffice.WordApi.Shape.LateBindingApiWrapperType, new object[]{ style, type, left, top, width, height, anchor });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Shape>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Shape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Shape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Shape>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Shape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Shape item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}