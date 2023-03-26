using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface Workbooks 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class Workbooks : COMObject, IEnumerableProvider<NetOffice.ExcelApi.Workbook>
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
                    _type = typeof(Workbooks);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Workbooks(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Workbooks(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(string progId) : base(progId)
		{
		}

        #endregion
        
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Application"/> </remarks>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Creator"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Parent"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Count"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.Workbook this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbook>(this, "_Default", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Add"/> </remarks>
		/// <param name="template">optional object template</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Add(object template)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Add", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, template);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Add", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Close"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Close()
		{
			 Factory.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="local">optional object local</param>
		/// <param name="corruptLoad">optional object corruptLoad</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local, object corruptLoad)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local, corruptLoad });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks, readOnly);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks, readOnly, format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.Open"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		/// <param name="local">optional object local</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename)
		{
			 Factory.ExecuteMethod(this, "_OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin)
		{
			 Factory.ExecuteMethod(this, "_OpenText", filename, origin);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow)
		{
			 Factory.ExecuteMethod(this, "_OpenText", filename, origin, startRow);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType)
		{
			 Factory.ExecuteMethod(this, "_OpenText", filename, origin, startRow, dataType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			 Factory.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
		/// <param name="local">optional object local</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers, object local)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename)
		{
			 Factory.ExecuteMethod(this, "OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin)
		{
			 Factory.ExecuteMethod(this, "OpenText", filename, origin);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow)
		{
			 Factory.ExecuteMethod(this, "OpenText", filename, origin, startRow);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType)
		{
			 Factory.ExecuteMethod(this, "OpenText", filename, origin, startRow, dataType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenText"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		/// <param name="decimalSeparator">optional object decimalSeparator</param>
		/// <param name="thousandsSeparator">optional object thousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object trailingMinusNumbers</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
		{
			 Factory.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		/// <param name="addToMru">optional object addToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks, readOnly);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, updateLinks, readOnly, format);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="delimiter">optional object delimiter</param>
		/// <param name="editable">optional object editable</param>
		/// <param name="notify">optional object notify</param>
		/// <param name="converter">optional object converter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		/// <param name="textVisualLayout">optional object textVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename)
		{
			 Factory.ExecuteMethod(this, "__OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin)
		{
			 Factory.ExecuteMethod(this, "__OpenText", filename, origin);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow)
		{
			 Factory.ExecuteMethod(this, "__OpenText", filename, origin, startRow);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType)
		{
			 Factory.ExecuteMethod(this, "__OpenText", filename, origin, startRow, dataType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		/// <param name="semicolon">optional object semicolon</param>
		/// <param name="comma">optional object comma</param>
		/// <param name="space">optional object space</param>
		/// <param name="other">optional object other</param>
		/// <param name="otherChar">optional object otherChar</param>
		/// <param name="fieldInfo">optional object fieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 Factory.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenDatabase"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="importDataAs">optional object importDataAs</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery, object importDataAs)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, new object[]{ filename, commandText, commandType, backgroundQuery, importDataAs });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenDatabase"/> </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenDatabase"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, commandText);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenDatabase"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, commandText, commandType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenDatabase"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, commandText, commandType, backgroundQuery);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.CheckOut"/> </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public void CheckOut(string filename)
		{
			 Factory.ExecuteMethod(this, "CheckOut", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.CanCheckOut"/> </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool CanCheckOut(string filename)
		{
			return Factory.ExecuteBoolMethodGet(this, "CanCheckOut", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenXML"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, stylesheets);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenXML"/> </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		/// <param name="loadOption">optional object loadOption</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets, object loadOption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, stylesheets, loadOption);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.Workbooks.OpenXML"/> </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _OpenXML(string filename, object stylesheets)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_OpenXML", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename, stylesheets);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _OpenXML(string filename)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_OpenXML", NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType, filename);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Workbook>

        ICOMObject IEnumerableProvider<Workbook>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<Workbook>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Workbook>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.Workbook> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Workbook item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

        #endregion

        #pragma warning restore
    }
}
