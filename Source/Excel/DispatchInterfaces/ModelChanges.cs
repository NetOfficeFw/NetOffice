using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface ModelChanges 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232191.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ModelChanges : COMObject
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
                    _type = typeof(ModelChanges);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ModelChanges(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ModelChanges(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229191.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230780.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230510.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231423.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesAdded
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelTableNames>(this, "TablesAdded", NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228365.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesDeleted
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelTableNames>(this, "TablesDeleted", NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231551.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNames TablesModified
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelTableNames>(this, "TablesModified", NetOffice.ExcelApi.ModelTableNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231611.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelTableNameChanges TableNamesChanged
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelTableNameChanges>(this, "TableNamesChanged", NetOffice.ExcelApi.ModelTableNameChanges.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232132.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool RelationshipChange
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RelationshipChange");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231460.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnNames ColumnsAdded
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelColumnNames>(this, "ColumnsAdded", NetOffice.ExcelApi.ModelColumnNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232085.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnNames ColumnsDeleted
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelColumnNames>(this, "ColumnsDeleted", NetOffice.ExcelApi.ModelColumnNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228803.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelColumnChanges ColumnsChanged
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelColumnChanges>(this, "ColumnsChanged", NetOffice.ExcelApi.ModelColumnChanges.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232121.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.ModelMeasureNames MeasuresAdded
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ModelMeasureNames>(this, "MeasuresAdded", NetOffice.ExcelApi.ModelMeasureNames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230632.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public bool UnknownChange
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UnknownChange");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/dn448395.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.Enums.XlModelChangeSource Source
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlModelChangeSource>(this, "Source");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
