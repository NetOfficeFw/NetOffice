using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Profile 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920658(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("6711F81E-AA55-47EE-AA9F-31B00E709CB6")]
	public interface Profile : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string Server { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjLoginType LoginType { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string UserName { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjProfileType Type { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjProfileConnectionState ConnectionState { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		string SiteId { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		void Delete();

		#endregion
	}
}
