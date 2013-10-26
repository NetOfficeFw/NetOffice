using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767247(v=office.14).aspx </remarks>
	[SupportByVersionAttribute("Visio", 14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisRoleSelectionTypes
	{
		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRoleSelConnector = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRoleSelContainer = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visRoleSelCallout = 4
	}
}