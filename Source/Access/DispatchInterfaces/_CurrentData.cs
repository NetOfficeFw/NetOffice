using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _CurrentData 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("9212BA73-3E79-11D1-98BD-006008197D41")]
    [CoClassSource(typeof(NetOffice.AccessApi.CodeData), typeof(NetOffice.AccessApi.CurrentData))]
    public interface _CurrentData : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196394.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.AllTables AllTables { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191708.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.AllQueries AllQueries { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837159.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.AllViews AllViews { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836247.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.AllStoredProcedures AllStoredProcedures { get; }

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834696.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		NetOffice.AccessApi.AllDatabaseDiagrams AllDatabaseDiagrams { get; }

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196686.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		NetOffice.AccessApi.AllFunctions AllFunctions { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		bool IsMemberSafe(Int32 dispid);

		#endregion
	}
}
