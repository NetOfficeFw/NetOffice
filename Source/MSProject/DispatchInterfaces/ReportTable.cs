using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface ReportTable 
	/// SupportByVersion MSProject, 11
	///</summary>
	[SupportByVersionAttribute("MSProject", 11)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ReportTable : COMObject
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
                    _type = typeof(ReportTable);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("MSProject", 11)]
		public Int32 RowsCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RowsCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11)]
		public Int32 ColumnsCount
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ColumnsCount", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="task">bool Task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		/// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
		/// <param name="safeArrayOfPjField">optional object SafeArrayOfPjField</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel, object safeArrayOfPjField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(task, groupName, filterName, outlineLevel, safeArrayOfPjField);
			Invoker.Method(this, "UpdateTableData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="task">bool Task</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public void UpdateTableData(bool task)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(task);
			Invoker.Method(this, "UpdateTableData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="task">bool Task</param>
		/// <param name="groupName">optional string GroupName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(task, groupName);
			Invoker.Method(this, "UpdateTableData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="task">bool Task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(task, groupName, filterName);
			Invoker.Method(this, "UpdateTableData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="task">bool Task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		/// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11)]
		public void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(task, groupName, filterName, outlineLevel);
			Invoker.Method(this, "UpdateTableData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// 
		/// </summary>
		/// <param name="row">Int32 Row</param>
		/// <param name="col">Int32 Col</param>
		[SupportByVersionAttribute("MSProject", 11)]
		public string GetCellText(Int32 row, Int32 col)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(row, col);
			object returnItem = Invoker.MethodReturn(this, "GetCellText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}