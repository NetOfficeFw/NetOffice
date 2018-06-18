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
	[TypeId("33DAA9FA-94CA-414E-BCF4-3E260B02B55E")]
	public interface ReportTable : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 RowsCount { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 ColumnsCount { get; }

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
		void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel, object safeArrayOfPjField);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void UpdateTableData(bool task);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void UpdateTableData(bool task, object groupName);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void UpdateTableData(bool task, object groupName, object filterName);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="task">bool task</param>
		/// <param name="groupName">optional string GroupName = </param>
		/// <param name="filterName">optional string FilterName = </param>
		/// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11)]
		void UpdateTableData(bool task, object groupName, object filterName, object outlineLevel);

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		/// <param name="row">Int32 row</param>
		/// <param name="col">Int32 col</param>
		[SupportByVersion("MSProject", 11)]
		string GetCellText(Int32 row, Int32 col);

		#endregion
	}
}
