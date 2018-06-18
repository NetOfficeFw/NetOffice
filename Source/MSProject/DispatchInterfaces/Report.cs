using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface Report 
	/// SupportByVersion MSProject, 11
	/// </summary>
	[SupportByVersion("MSProject", 11)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("C4F74DB0-E706-4AD5-9E71-AD5E6A0CF388")]
	public interface Report : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		string Name { get; set; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		Int32 Index { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Project Parent { get; }

		/// <summary>
		/// SupportByVersion MSProject 11
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		NetOffice.MSProjectApi.Shapes Shapes { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Delete();

		/// <summary>
		/// SupportByVersion MSProject 11
		/// </summary>
		[SupportByVersion("MSProject", 11)]
		void Apply();

		#endregion
	}
}
