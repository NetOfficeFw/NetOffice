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
	/// DispatchInterface Documents 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.documents"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Documents : COMObject, IEnumerableProvider<NetOffice.WordApi.Document>
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
                    _type = typeof(Documents);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Documents(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Documents(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Documents(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
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
		public NetOffice.WordApi.Document this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Item", NetOffice.WordApi.Document.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Close"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		/// <param name="routeDocument">optional object routeDocument</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat, object routeDocument)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, originalFormat, routeDocument);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Close"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Close"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Close"/> </remarks>
		/// <param name="saveChanges">optional object saveChanges</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Close(object saveChanges, object originalFormat)
		{
			 Factory.ExecuteMethod(this, "Close", saveChanges, originalFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld(object template, object newTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "AddOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, template, newTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "AddOld", NetOffice.WordApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="template">optional object template</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document AddOld(object template)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "AddOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, template);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document OpenOld(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenOld", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Save"/> </remarks>
		/// <param name="noPrompt">optional object noPrompt</param>
		/// <param name="originalFormat">optional object originalFormat</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Save(object noPrompt, object originalFormat)
		{
			 Factory.ExecuteMethod(this, "Save", noPrompt, originalFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Save"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Save"/> </remarks>
		/// <param name="noPrompt">optional object noPrompt</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Save(object noPrompt)
		{
			 Factory.ExecuteMethod(this, "Save", noPrompt);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Add"/> </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		/// <param name="documentType">optional object documentType</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Add", NetOffice.WordApi.Document.LateBindingApiWrapperType, template, newTemplate, documentType, visible);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Add", NetOffice.WordApi.Document.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Add"/> </remarks>
		/// <param name="template">optional object template</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Add", NetOffice.WordApi.Document.LateBindingApiWrapperType, template);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Add"/> </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Add", NetOffice.WordApi.Document.LateBindingApiWrapperType, template, newTemplate);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Add"/> </remarks>
		/// <param name="template">optional object template</param>
		/// <param name="newTemplate">optional object newTemplate</param>
		/// <param name="documentType">optional object documentType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Add(object template, object newTemplate, object documentType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Add", NetOffice.WordApi.Document.LateBindingApiWrapperType, template, newTemplate, documentType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		/// <param name="xMLTransform">optional object xMLTransform</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog, xMLTransform });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.Open"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2000(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2000", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.CheckOut"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void CheckOut(string fileName)
		{
			 Factory.ExecuteMethod(this, "CheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.CanCheckOut"/> </remarks>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool CanCheckOut(string fileName)
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckOut", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Document Open2002(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "Open2002", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		/// <param name="xMLTransform">optional object xMLTransform</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog, object xMLTransform)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog, xMLTransform });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, fileName, confirmConversions, readOnly, addToRecentFiles);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.OpenNoRepairDialog"/> </remarks>
		/// <param name="fileName">object fileName</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="addToRecentFiles">optional object addToRecentFiles</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="revert">optional object revert</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="format">optional object format</param>
		/// <param name="encoding">optional object encoding</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="openAndRepair">optional object openAndRepair</param>
		/// <param name="documentDirection">optional object documentDirection</param>
		/// <param name="noEncodingDialog">optional object noEncodingDialog</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document OpenNoRepairDialog(object fileName, object confirmConversions, object readOnly, object addToRecentFiles, object passwordDocument, object passwordTemplate, object revert, object writePasswordDocument, object writePasswordTemplate, object format, object encoding, object visible, object openAndRepair, object documentDirection, object noEncodingDialog)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "OpenNoRepairDialog", NetOffice.WordApi.Document.LateBindingApiWrapperType, new object[]{ fileName, confirmConversions, readOnly, addToRecentFiles, passwordDocument, passwordTemplate, revert, writePasswordDocument, writePasswordTemplate, format, encoding, visible, openAndRepair, documentDirection, noEncodingDialog });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.AddBlogDocument"/> </remarks>
		/// <param name="providerID">string providerID</param>
		/// <param name="postURL">string postURL</param>
		/// <param name="blogName">string blogName</param>
		/// <param name="postID">optional string PostID = </param>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName, object postID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "AddBlogDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, providerID, postURL, blogName, postID);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Documents.AddBlogDocument"/> </remarks>
		/// <param name="providerID">string providerID</param>
		/// <param name="postURL">string postURL</param>
		/// <param name="blogName">string blogName</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Document AddBlogDocument(string providerID, string postURL, string blogName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Document>(this, "AddBlogDocument", NetOffice.WordApi.Document.LateBindingApiWrapperType, providerID, postURL, blogName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Document>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Document>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Document>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion
        
        #region IEnumerable<NetOffice.WordApi.Document>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Document> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Document item in innerEnumerator)
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