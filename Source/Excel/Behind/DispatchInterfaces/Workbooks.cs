using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface Workbooks 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841074.aspx </remarks>
	public class Workbooks : COMObject, NetOffice.ExcelApi.Workbooks
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ExcelApi.Workbooks);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Workbooks() : base()
		{

		}

		#endregion
        
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195019.aspx </remarks>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195436.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837124.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822893.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.ExcelApi.Workbook this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Workbook>(this, "_Default", typeof(NetOffice.ExcelApi.Workbook), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx </remarks>
		/// <param name="template">optional object template</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Add(object template)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Add", typeof(NetOffice.ExcelApi.Workbook), template);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Add()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Add", typeof(NetOffice.ExcelApi.Workbook));
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839657.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void Close()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close");
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local, object corruptLoad)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local, corruptLoad });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks, readOnly);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks, readOnly, format);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		/// <param name="readOnly">optional object readOnly</param>
		/// <param name="format">optional object format</param>
		/// <param name="password">optional object password</param>
		/// <param name="writeResPassword">optional object writeResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object ignoreReadOnlyRecommended</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx </remarks>
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
		public virtual NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void _OpenText(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void _OpenText(string filename, object origin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", filename, origin);
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
		public virtual void _OpenText(string filename, object origin, object startRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", filename, origin, startRow);
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", filename, origin, startRow, dataType);
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
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
		public virtual void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "_OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers, object local)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers, local });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", filename, origin);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin, object startRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", filename, origin, startRow);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin, object startRow, object dataType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", filename, origin, startRow, dataType);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		/// <param name="startRow">optional object startRow</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object consecutiveDelimiter</param>
		/// <param name="tab">optional object tab</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx </remarks>
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
		public virtual void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="updateLinks">optional object updateLinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks);
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks, readOnly);
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), filename, updateLinks, readOnly, format);
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify });
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
		public virtual NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_Open", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void __OpenText(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="origin">optional object origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void __OpenText(string filename, object origin)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", filename, origin);
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
		public virtual void __OpenText(string filename, object origin, object startRow)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", filename, origin, startRow);
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", filename, origin, startRow, dataType);
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar });
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
		public virtual void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "__OpenText", new object[]{ filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		/// <param name="importDataAs">optional object importDataAs</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery, object importDataAs)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", typeof(NetOffice.ExcelApi.Workbook), new object[]{ filename, commandText, commandType, backgroundQuery, importDataAs });
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenDatabase(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", typeof(NetOffice.ExcelApi.Workbook), filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", typeof(NetOffice.ExcelApi.Workbook), filename, commandText);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", typeof(NetOffice.ExcelApi.Workbook), filename, commandText, commandType);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="commandText">optional object commandText</param>
		/// <param name="commandType">optional object commandType</param>
		/// <param name="backgroundQuery">optional object backgroundQuery</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenDatabase", typeof(NetOffice.ExcelApi.Workbook), filename, commandText, commandType, backgroundQuery);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194062.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void CheckOut(string filename)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CheckOut", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193284.aspx </remarks>
		/// <param name="filename">string filename</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool CanCheckOut(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "CanCheckOut", filename);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", typeof(NetOffice.ExcelApi.Workbook), filename, stylesheets);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		/// <param name="loadOption">optional object loadOption</param>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets, object loadOption)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", typeof(NetOffice.ExcelApi.Workbook), filename, stylesheets, loadOption);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx </remarks>
		/// <param name="filename">string filename</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook OpenXML(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "OpenXML", typeof(NetOffice.ExcelApi.Workbook), filename);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		/// <param name="stylesheets">optional object stylesheets</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook _OpenXML(string filename, object stylesheets)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_OpenXML", typeof(NetOffice.ExcelApi.Workbook), filename, stylesheets);
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="filename">string filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Workbook _OpenXML(string filename)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Workbook>(this, "_OpenXML", typeof(NetOffice.ExcelApi.Workbook), filename);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Workbook>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Workbook>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Workbook>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Workbook>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.Workbook> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

        #endregion

        #pragma warning restore
    }
}


