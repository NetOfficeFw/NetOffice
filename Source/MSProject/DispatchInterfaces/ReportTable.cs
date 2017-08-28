using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface ReportTable 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ReportTable : COMObject
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
                    _type = typeof(ReportTable);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ReportTable(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ReportTable(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ReportTable(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 RowsCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "RowsCount");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public Int32 ColumnsCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColumnsCount");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		/// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
		/// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
		[SupportByVersion("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel, object safeArrayOfPjField)
		{
			 Factory.ExecuteMethod(this, "UpdateTableData", new object[]{ task, groupName, filterName, outlineLevel, safeArrayOfPjField });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public void UpdateTableData(bool task)
		{
			 Factory.ExecuteMethod(this, "UpdateTableData", task);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName)
		{
			 Factory.ExecuteMethod(this, "UpdateTableData", task, groupName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName)
		{
			 Factory.ExecuteMethod(this, "UpdateTableData", task, groupName, filterName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		/// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel)
		{
			 Factory.ExecuteMethod(this, "UpdateTableData", task, groupName, filterName, outlineLevel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="col">Int32 col</param>
		[SupportByVersion("MSProject", 11)]
		public string GetCellText(Int32 row, Int32 col)
		{
			return Factory.ExecuteStringMethodGet(this, "GetCellText", row, col);
		}

		#endregion

		#pragma warning restore
	}
}
