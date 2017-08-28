using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 12, 14, 15, 16
	 /// </summary>
	[SupportByVersion("Visio", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum VisAutoConnectDir
	{
		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoConnectDirNone = 0,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoConnectDirUp = 1,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoConnectDirDown = 2,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoConnectDirLeft = 3,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Visio", 12,14,15,16)]
		 visAutoConnectDirRight = 4
	}
}