using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface CubeField 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822619.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
	public interface CubeField : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838570.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195332.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822639.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835244.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCubeFieldType CubeFieldType { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196833.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		string Caption { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837351.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		string Name { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198103.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		string Value { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837998.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlPivotFieldOrientation Orientation { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841272.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Position { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820886.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.TreeviewControl TreeviewControl { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194397.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool DragToColumn { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820897.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool DragToHide { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836521.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool DragToPage { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193981.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool DragToRow { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198094.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		bool DragToData { get; set; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 HiddenLevels { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838439.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool HasMemberProperties { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837592.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlLayoutFormType LayoutForm { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839474.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.PivotFields PivotFields { get; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835260.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool EnableMultiplePageItems { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838068.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.xLSubtototalLocationType LayoutSubtotalLocation { get; set; }

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836523.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		bool ShowInFieldList { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		string _Caption { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834624.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		bool IncludeNewItemsInFilter { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839077.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCubeFieldSubType CubeFieldSubType { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836450.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		bool AllItemsVisible { get; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841249.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		string CurrentPageName { get; set; }

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193506.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		bool IsDate { get; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838407.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool FlattenHierarchies { get; set; }

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822855.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		bool HierarchizeDistinct { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822896.aspx </remarks>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void AddMemberPropertyField(string property, object propertyOrder);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822896.aspx </remarks>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		/// <param name="propertyDisplayedIn">optional object propertyDisplayedIn</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		void AddMemberPropertyField(string property, object propertyOrder, object propertyDisplayedIn);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822896.aspx </remarks>
		/// <param name="property">string property</param>
		[CustomMethod]
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void AddMemberPropertyField(string property);

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838430.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		void Delete();

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="property">string property</param>
		/// <param name="propertyOrder">optional object propertyOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 12,14,15,16)]
		void _AddMemberPropertyField(string property, object propertyOrder);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="property">string property</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		void _AddMemberPropertyField(string property);

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195930.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		void ClearManualFilter();

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835266.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		void CreatePivotFields();

		#endregion
	}
}
