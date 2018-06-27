using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSProjectApi;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// DispatchInterface ReportTable 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ReportTable : COMObject, NetOffice.MSProjectApi.ReportTable
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
                    _contractType = typeof(NetOffice.MSProjectApi.ReportTable);
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
                    _type = typeof(ReportTable);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ReportTable() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual Int32 RowsCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RowsCount");
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		public virtual Int32 ColumnsCount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ColumnsCount");
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
		public virtual void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel, object safeArrayOfPjField)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateTableData", new object[]{ task, groupName, filterName, outlineLevel, safeArrayOfPjField });
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public virtual void UpdateTableData(bool task)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateTableData", task);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public virtual void UpdateTableData(bool task, object groupName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateTableData", task, groupName);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		public virtual void UpdateTableData(bool task, object groupName, object filterName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateTableData", task, groupName, filterName);
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
		public virtual void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateTableData", task, groupName, filterName, outlineLevel);
		}

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="col">Int32 col</param>
		[SupportByVersion("MSProject", 11)]
		public virtual string GetCellText(Int32 row, Int32 col)
		{
			return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "GetCellText", row, col);
		}

		#endregion

		#pragma warning restore
	}
}

