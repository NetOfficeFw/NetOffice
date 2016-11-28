using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// DispatchInterface ModelChanges 
	/// SupportByVersion Excel, 15, 16
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232191.aspx
	///</summary>
	[SupportByVersionAttribute("Excel", 15, 16)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ModelChanges : COMObject
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
                    _type = typeof(ModelChanges);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ModelChanges(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229191.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
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
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230780.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
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
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230510.aspx
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
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
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231423.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesAdded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesAdded", paramsArray);
				NetOffice.ExcelApi.ModelTableNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelTableNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228365.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesDeleted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesDeleted", paramsArray);
				NetOffice.ExcelApi.ModelTableNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelTableNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231551.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesModified
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TablesModified", paramsArray);
				NetOffice.ExcelApi.ModelTableNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelTableNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231611.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNameChanges TableNamesChanged
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "TableNamesChanged", paramsArray);
				NetOffice.ExcelApi.ModelTableNameChanges newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelTableNameChanges.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelTableNameChanges;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232132.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool RelationshipChange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RelationshipChange", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj231460.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnNames ColumnsAdded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnsAdded", paramsArray);
				NetOffice.ExcelApi.ModelColumnNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelColumnNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelColumnNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232085.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnNames ColumnsDeleted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnsDeleted", paramsArray);
				NetOffice.ExcelApi.ModelColumnNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelColumnNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelColumnNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj228803.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnChanges ColumnsChanged
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnsChanged", paramsArray);
				NetOffice.ExcelApi.ModelColumnChanges newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelColumnChanges.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelColumnChanges;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj232121.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelMeasureNames MeasuresAdded
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MeasuresAdded", paramsArray);
				NetOffice.ExcelApi.ModelMeasureNames newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.ModelMeasureNames.LateBindingApiWrapperType) as NetOffice.ExcelApi.ModelMeasureNames;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj230632.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public bool UnknownChange
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UnknownChange", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448395.aspx
		/// </summary>
		[SupportByVersionAttribute("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlModelChangeSource Source
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Source", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlModelChangeSource)intReturnItem;
			}
		}

		#endregion

		#region Methods

		#endregion
		#pragma warning restore
	}
}