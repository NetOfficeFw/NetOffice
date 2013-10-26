using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisDataRecordsetAddOptions
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataRecordsetNoExternalDataUI = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataRecordsetNoRefreshUI = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataRecordsetNoAdvConfig = 4,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataRecordsetDelayQuery = 8,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visDataRecordsetDontCopyLinks = 16
	}
}