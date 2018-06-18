using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Group2 
	/// SupportByVersion MSProject, 11,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920602(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("11589058-69F0-11D2-B889-00C04FB90729")]
	public interface Group2 : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		Int32 Index { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.Project Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		NetOffice.MSProjectApi.GroupCriteria2 GroupCriteria { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool ShowSummary { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool GroupAssignments { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		bool MaintainHierarchy { get; set; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,14)]
		void Delete();

		#endregion
	}
}
