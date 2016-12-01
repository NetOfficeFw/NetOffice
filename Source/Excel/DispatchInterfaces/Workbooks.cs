using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// DispatchInterface Workbooks 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff841074.aspx
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Workbooks : COMObject ,IEnumerable<NetOffice.ExcelApi.Workbook>
	{
		#pragma warning disable
		#region Type Information

        /// <summary>
        /// Instance Type
        /// </summary>
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
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Workbooks(string progId) : base(progId)
		{
		}

        #endregion
        
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195019.aspx
        /// </summary>
        [SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195436.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837124.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822893.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.ExcelApi.Workbook this[object index]
		{
			get
            {			
			    object[] paramsArray = Invoker.ValidateParamsArray(index);
			    object returnItem = Invoker.PropertyGet(this, "_Default", paramsArray);
			    NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			    return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx
		/// </summary>
		/// <param name="template">optional object Template</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Add(object template)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(template);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840478.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839657.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void Close()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Close", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="local">optional object Local</param>
		/// <param name="corruptLoad">optional object CorruptLoad</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local, object corruptLoad)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local, corruptLoad);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194819.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		/// <param name="addToMru">optional object AddToMru</param>
		/// <param name="local">optional object Local</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru, local);
			object returnItem = Invoker.MethodReturn(this, "Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void _OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator);
			Invoker.Method(this, "_OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object TrailingMinusNumbers</param>
		/// <param name="local">optional object Local</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers, object local)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers, local);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff837097.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		/// <param name="decimalSeparator">optional object DecimalSeparator</param>
		/// <param name="thousandsSeparator">optional object ThousandsSeparator</param>
		/// <param name="trailingMinusNumbers">optional object TrailingMinusNumbers</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout, object decimalSeparator, object thousandsSeparator, object trailingMinusNumbers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout, decimalSeparator, thousandsSeparator, trailingMinusNumbers);
			Invoker.Method(this, "OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		/// <param name="addToMru">optional object AddToMru</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter, addToMru);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="updateLinks">optional object UpdateLinks</param>
		/// <param name="readOnly">optional object ReadOnly</param>
		/// <param name="format">optional object Format</param>
		/// <param name="password">optional object Password</param>
		/// <param name="writeResPassword">optional object WriteResPassword</param>
		/// <param name="ignoreReadOnlyRecommended">optional object IgnoreReadOnlyRecommended</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="delimiter">optional object Delimiter</param>
		/// <param name="editable">optional object Editable</param>
		/// <param name="notify">optional object Notify</param>
		/// <param name="converter">optional object Converter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _Open(string filename, object updateLinks, object readOnly, object format, object password, object writeResPassword, object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, updateLinks, readOnly, format, password, writeResPassword, ignoreReadOnlyRecommended, origin, delimiter, editable, notify, converter);
			object returnItem = Invoker.MethodReturn(this, "_Open", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		/// <param name="textVisualLayout">optional object TextVisualLayout</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo, object textVisualLayout)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo, textVisualLayout);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="origin">optional object Origin</param>
		/// <param name="startRow">optional object StartRow</param>
		/// <param name="dataType">optional object DataType</param>
		/// <param name="textQualifier">optional NetOffice.ExcelApi.Enums.XlTextQualifier TextQualifier = 1</param>
		/// <param name="consecutiveDelimiter">optional object ConsecutiveDelimiter</param>
		/// <param name="tab">optional object Tab</param>
		/// <param name="semicolon">optional object Semicolon</param>
		/// <param name="comma">optional object Comma</param>
		/// <param name="space">optional object Space</param>
		/// <param name="other">optional object Other</param>
		/// <param name="otherChar">optional object OtherChar</param>
		/// <param name="fieldInfo">optional object FieldInfo</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void __OpenText(string filename, object origin, object startRow, object dataType, object textQualifier, object consecutiveDelimiter, object tab, object semicolon, object comma, object space, object other, object otherChar, object fieldInfo)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, origin, startRow, dataType, textQualifier, consecutiveDelimiter, tab, semicolon, comma, space, other, otherChar, fieldInfo);
			Invoker.Method(this, "__OpenText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="commandText">optional object CommandText</param>
		/// <param name="commandType">optional object CommandType</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		/// <param name="importDataAs">optional object ImportDataAs</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery, object importDataAs)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, commandText, commandType, backgroundQuery, importDataAs);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="commandText">optional object CommandText</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, commandText);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="commandText">optional object CommandText</param>
		/// <param name="commandType">optional object CommandType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, commandText, commandType);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193543.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="commandText">optional object CommandText</param>
		/// <param name="commandType">optional object CommandType</param>
		/// <param name="backgroundQuery">optional object BackgroundQuery</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenDatabase(string filename, object commandText, object commandType, object backgroundQuery)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, commandText, commandType, backgroundQuery);
			object returnItem = Invoker.MethodReturn(this, "OpenDatabase", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194062.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public void CheckOut(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			Invoker.Method(this, "CheckOut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193284.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public bool CanCheckOut(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "CanCheckOut", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="stylesheets">optional object Stylesheets</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, stylesheets);
			object returnItem = Invoker.MethodReturn(this, "OpenXML", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="stylesheets">optional object Stylesheets</param>
		/// <param name="loadOption">optional object LoadOption</param>
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename, object stylesheets, object loadOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, stylesheets, loadOption);
			object returnItem = Invoker.MethodReturn(this, "OpenXML", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838643.aspx
		/// </summary>
		/// <param name="filename">string Filename</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook OpenXML(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "OpenXML", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		/// <param name="stylesheets">optional object Stylesheets</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _OpenXML(string filename, object stylesheets)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename, stylesheets);
			object returnItem = Invoker.MethodReturn(this, "_OpenXML", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="filename">string Filename</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Workbook _OpenXML(string filename)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(filename);
			object returnItem = Invoker.MethodReturn(this, "_OpenXML", paramsArray);
			NetOffice.ExcelApi.Workbook newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.Workbook.LateBindingApiWrapperType) as NetOffice.ExcelApi.Workbook;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.ExcelApi.Workbook> Member
        
        /// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
       public IEnumerator<NetOffice.ExcelApi.Workbook> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.ExcelApi.Workbook item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15,16
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}