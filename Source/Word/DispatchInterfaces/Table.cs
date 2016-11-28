using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface Table 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834860.aspx
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Table : COMObject
	{
		#pragma warning disable
		#region Type Information

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Table);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Table(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Table(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197195.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Range", paramsArray);
				NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839082.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845545.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835743.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198160.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Columns Columns
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Columns", paramsArray);
				NetOffice.WordApi.Columns newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Columns.LateBindingApiWrapperType) as NetOffice.WordApi.Columns;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839587.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Rows Rows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Rows", paramsArray);
				NetOffice.WordApi.Rows newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Rows.LateBindingApiWrapperType) as NetOffice.WordApi.Rows;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff823239.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Borders Borders
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Borders", paramsArray);
				NetOffice.WordApi.Borders newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Borders.LateBindingApiWrapperType) as NetOffice.WordApi.Borders;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Borders", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845404.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shading Shading
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Shading", paramsArray);
				NetOffice.WordApi.Shading newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Shading.LateBindingApiWrapperType) as NetOffice.WordApi.Shading;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835471.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool Uniform
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Uniform", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193447.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 AutoFormatType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AutoFormatType", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836124.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Tables", paramsArray);
				NetOffice.WordApi.Tables newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Tables.LateBindingApiWrapperType) as NetOffice.WordApi.Tables;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194409.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Int32 NestingLevel
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "NestingLevel", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool AllowPageBreaks
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllowPageBreaks", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AllowPageBreaks", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839810.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public bool AllowAutoFit
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllowAutoFit", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "AllowAutoFit", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845887.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single PreferredWidth
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreferredWidth", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PreferredWidth", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834288.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPreferredWidthType PreferredWidthType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PreferredWidthType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdPreferredWidthType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PreferredWidthType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844783.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single TopPadding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TopPadding", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TopPadding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838742.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single BottomPadding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BottomPadding", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "BottomPadding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836311.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single LeftPadding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LeftPadding", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "LeftPadding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835709.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single RightPadding
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RightPadding", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RightPadding", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196121.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public Single Spacing
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Spacing", paramsArray);
				return NetRuntimeSystem.Convert.ToSingle(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Spacing", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193098.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTableDirection TableDirection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableDirection", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdTableDirection)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "TableDirection", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff840350.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public string ID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196552.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public object Style
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Style", paramsArray);
				if((null != returnItem) && (returnItem is MarshalByRefObject))
				{
					ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
					return newObject;
				}
				else
				{
					return  returnItem;
				}
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Style", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff191959.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleHeadingRows
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleHeadingRows", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleHeadingRows", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844792.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleLastRow
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleLastRow", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleLastRow", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff834832.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleFirstColumn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleFirstColumn", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleFirstColumn", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839138.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14,15,16)]
		public bool ApplyStyleLastColumn
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleLastColumn", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleLastColumn", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192619.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool ApplyStyleRowBands
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleRowBands", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleRowBands", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839106.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public bool ApplyStyleColumnBands
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ApplyStyleColumnBands", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ApplyStyleColumnBands", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835972.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public string Title
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Title", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Title", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820918.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 14,15,16)]
		public string Descr
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Descr", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Descr", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194359.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845868.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="languageID">optional object LanguageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, languageID);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive);
			Invoker.Method(this, "SortOld", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196507.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortAscending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortAscending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff835818.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void SortDescending()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SortDescending", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		/// <param name="autoFit">optional object AutoFit</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838081.aspx
		/// </summary>
		/// <param name="format">optional object Format</param>
		/// <param name="applyBorders">optional object ApplyBorders</param>
		/// <param name="applyShading">optional object ApplyShading</param>
		/// <param name="applyFont">optional object ApplyFont</param>
		/// <param name="applyColor">optional object ApplyColor</param>
		/// <param name="applyHeadingRows">optional object ApplyHeadingRows</param>
		/// <param name="applyLastRow">optional object ApplyLastRow</param>
		/// <param name="applyFirstColumn">optional object ApplyFirstColumn</param>
		/// <param name="applyLastColumn">optional object ApplyLastColumn</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat(object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn);
			Invoker.Method(this, "AutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838712.aspx
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void UpdateAutoFormat()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateAutoFormat", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToTextOld(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToTextOld", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToTextOld()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToTextOld", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff821612.aspx
		/// </summary>
		/// <param name="row">Int32 Row</param>
		/// <param name="column">Int32 Column</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Cell Cell(Int32 row, Int32 column)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(row, column);
			object returnItem = Invoker.MethodReturn(this, "Cell", paramsArray);
			NetOffice.WordApi.Cell newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Cell.LateBindingApiWrapperType) as NetOffice.WordApi.Cell;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836035.aspx
		/// </summary>
		/// <param name="beforeRow">object BeforeRow</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table Split(object beforeRow)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beforeRow);
			object returnItem = Invoker.MethodReturn(this, "Split", paramsArray);
			NetOffice.WordApi.Table newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Table.LateBindingApiWrapperType) as NetOffice.WordApi.Table;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		/// <param name="nestedTables">optional object NestedTables</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText(object separator, object nestedTables)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator, nestedTables);
			object returnItem = Invoker.MethodReturn(this, "ConvertToText", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ConvertToText", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820974.aspx
		/// </summary>
		/// <param name="separator">optional object Separator</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range ConvertToText(object separator)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(separator);
			object returnItem = Invoker.MethodReturn(this, "ConvertToText", paramsArray);
			NetOffice.WordApi.Range newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.Range.LateBindingApiWrapperType) as NetOffice.WordApi.Range;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff820953.aspx
		/// </summary>
		/// <param name="behavior">NetOffice.WordApi.Enums.WdAutoFitBehavior Behavior</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void AutoFitBehavior(NetOffice.WordApi.Enums.WdAutoFitBehavior behavior)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(behavior);
			Invoker.Method(this, "AutoFitBehavior", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		/// <param name="languageID">optional object LanguageID</param>
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192792.aspx
		/// </summary>
		/// <param name="excludeHeader">optional object ExcludeHeader</param>
		/// <param name="fieldNumber">optional object FieldNumber</param>
		/// <param name="sortFieldType">optional object SortFieldType</param>
		/// <param name="sortOrder">optional object SortOrder</param>
		/// <param name="fieldNumber2">optional object FieldNumber2</param>
		/// <param name="sortFieldType2">optional object SortFieldType2</param>
		/// <param name="sortOrder2">optional object SortOrder2</param>
		/// <param name="fieldNumber3">optional object FieldNumber3</param>
		/// <param name="sortFieldType3">optional object SortFieldType3</param>
		/// <param name="sortOrder3">optional object SortOrder3</param>
		/// <param name="caseSensitive">optional object CaseSensitive</param>
		/// <param name="bidiSort">optional object BidiSort</param>
		/// <param name="ignoreThe">optional object IgnoreThe</param>
		/// <param name="ignoreKashida">optional object IgnoreKashida</param>
		/// <param name="ignoreDiacritics">optional object IgnoreDiacritics</param>
		/// <param name="ignoreHe">optional object IgnoreHe</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192363.aspx
		/// </summary>
		/// <param name="styleName">string StyleName</param>
		[SupportByVersionAttribute("Word", 12,14,15,16)]
		public void ApplyStyleDirectFormatting(string styleName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(styleName);
			Invoker.Method(this, "ApplyStyleDirectFormatting", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}