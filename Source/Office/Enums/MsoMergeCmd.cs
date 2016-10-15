using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 15,16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227893.aspx </remarks>
	[SupportByVersionAttribute("Office", 15, 16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoMergeCmd
	{
		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoMergeUnion = 1,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoMergeCombine = 2,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoMergeIntersect = 3,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoMergeSubtract = 4,

		 /// <summary>
		 /// SupportByVersion Office 15,16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 15, 16)]
		 msoMergeFragment = 5
	}
}