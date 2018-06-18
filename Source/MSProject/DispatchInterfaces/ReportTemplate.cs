using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface ReportTemplate 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920672(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
	[TypeId("46BFF766-F07A-409C-A85F-EED0A36589F9")]
	public interface ReportTemplate : ICOMObject
	{
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		string TemplatePath { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjVisualReportsCubeType CubeType { get; }

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		NetOffice.MSProjectApi.Enums.PjVisualReportsTemplateType TemplateType { get; }

		#endregion

	}
}
