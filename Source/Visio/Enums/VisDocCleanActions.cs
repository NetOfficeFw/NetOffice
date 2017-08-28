using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisDocCleanActions
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActLocalFormulas = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActEmptyRowsAndSects = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActNonDefaultFonts = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActStaleResults = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActMissingSubs = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActConstantFormulas = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActNearZero = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActDuplicateSubs = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActBadDisplayLists = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActBadFieldCounts = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActDeletedFields = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActBadFieldFormulas = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActBadFieldMarks = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActBadReferences = 8192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16383</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActAll = 16383,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8152</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanActDefault = 8152,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanAlertDefault = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>984</remarks>
		 [SupportByVersion("Visio", 11,12,14,15,16)]
		 visDocCleanFixDefault = 984
	}
}