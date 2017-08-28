using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsEnum)]
	public enum FieldListObjectTypeEnum
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flTables = 1,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flViews = 2,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flStoredProcedures = 4,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flCmdText = 8,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flSchemaDiagrams = 16,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flOLAPCube = 32,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("OWC10", 1)]
		 flAll = 63
	}
}