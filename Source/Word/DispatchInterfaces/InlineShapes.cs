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
	/// DispatchInterface InlineShapes 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class InlineShapes : COMObject, IEnumerableProvider<NetOffice.WordApi.InlineShape>
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
                    _type = typeof(InlineShapes);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public InlineShapes(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public InlineShapes(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public InlineShapes(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.Count"/> </remarks>
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
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.InlineShape this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "Item", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile, object saveWithDocument, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPicture", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPicture", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPicture", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName, linkToFile);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPicture"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="saveWithDocument">optional object saveWithDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPicture(string fileName, object linkToFile, object saveWithDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPicture", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName, linkToFile, saveWithDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel, range });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType, fileName, linkToFile);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType, fileName, linkToFile, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEObject"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="fileName">optional object fileName</param>
		/// <param name="linkToFile">optional object linkToFile</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEObject(object classType, object fileName, object linkToFile, object displayAsIcon, object iconFileName, object iconIndex, object iconLabel)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEObject", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ classType, fileName, linkToFile, displayAsIcon, iconFileName, iconIndex, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEControl(object classType, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEControl", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEControl"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEControl()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEControl", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddOLEControl"/> </remarks>
		/// <param name="classType">optional object classType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddOLEControl(object classType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddOLEControl", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, classType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.New"/> </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape New(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "New", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddHorizontalLine"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddHorizontalLine(string fileName, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddHorizontalLine", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddHorizontalLine"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddHorizontalLine(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddHorizontalLine", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddHorizontalLineStandard"/> </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddHorizontalLineStandard(object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddHorizontalLineStandard", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddHorizontalLineStandard"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddHorizontalLineStandard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddHorizontalLineStandard", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPictureBullet"/> </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPictureBullet(string fileName, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPictureBullet", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddPictureBullet"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddPictureBullet(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddPictureBullet", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddChart(object type, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, type, range);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddChart()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.InlineShape AddChart(object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.InlineShape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddSmartArt", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, layout, range);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.InlineShapes.AddSmartArt"/> </remarks>
		/// <param name="layout">NetOffice.OfficeApi.SmartArtLayout layout</param>
		[CustomMethod]
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.InlineShape AddSmartArt(NetOffice.OfficeApi.SmartArtLayout layout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddSmartArt", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, layout);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddWebVideo", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url, range });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddWebVideo(string embedCode, object videoWidth, object videoHeight)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddWebVideo", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, embedCode, videoWidth, videoHeight);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddWebVideo", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, embedCode, videoWidth, videoHeight, posterFrameImage);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addwebvideo"/> </remarks>
		/// <param name="embedCode">string embedCode</param>
		/// <param name="videoWidth">object videoWidth</param>
		/// <param name="videoHeight">object videoHeight</param>
		/// <param name="posterFrameImage">optional object posterFrameImage</param>
		/// <param name="url">optional object url</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddWebVideo(string embedCode, object videoWidth, object videoHeight, object posterFrameImage, object url)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddWebVideo", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, new object[]{ embedCode, videoWidth, videoHeight, posterFrameImage, url });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="range">optional object range</param>
		/// <param name="newLayout">optional object newLayout</param>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddChart2(object style, object type, object range, object newLayout)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart2", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, style, type, range, newLayout);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addchart2"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddChart2()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart2", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddChart2(object style)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart2", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, style);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddChart2(object style, object type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart2", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, style, type);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.inlineshapes.addchart2"/> </remarks>
		/// <param name="style">optional Int32 Style = -1</param>
		/// <param name="type">optional NetOffice.OfficeApi.Enums.XlChartType Type = -1</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.InlineShape AddChart2(object style, object type, object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.InlineShape>(this, "AddChart2", NetOffice.WordApi.InlineShape.LateBindingApiWrapperType, style, type, range);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.InlineShape>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.InlineShape>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.InlineShape>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.InlineShape>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.InlineShape> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.InlineShape item in innerEnumerator)
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