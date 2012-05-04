using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisContainerMemberState
	{
		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visContainerMemberNotAMember = 0,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visContainerMemberInterior = 1,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visContainerMemberOnBoundary = 2,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visContainerMemberOutside = 3,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visContainerMemberInList = 4
	}
}