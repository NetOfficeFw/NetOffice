using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisDefaultSaveFormats
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visDefaultSaveCurrentBinary = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visDefaultSavePreviousBinary = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visDefaultSaveCurrentXML = 2,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visDefaultSaveCurrent = 0,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visDefaultSaveCurrentMacroEnabled = 3
	}
}