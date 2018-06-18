using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface ViewsSingle 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920756(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("B4097887-EC12-4683-9622-9974EF26353C")]
	public interface ViewsSingle : Views
	{
		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		/// <param name="highlightFilt">optional bool HighlightFilt = false</param>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group, object highlightFilt);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter);

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group);

		#endregion
	}
}
