using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/overview/Library-Reference/msomergecmd-enumeration-office"/> </remarks>
	[SupportByVersion("Office", 15, 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoMergeCmd
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoMergeUnion = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoMergeCombine = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoMergeIntersect = 3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoMergeSubtract = 4,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 15, 16)]
		 msoMergeFragment = 5
	}
}