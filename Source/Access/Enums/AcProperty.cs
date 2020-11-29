using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AcProperty"/> </remarks>
	[SupportByVersion("Access", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcProperty
	{
		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyEnabled = 0,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyVisible = 1,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyLocked = 2,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyLeft = 3,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyTop = 4,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyWidth = 5,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyHeight = 6,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyForeColor = 7,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyBackColor = 8,

		 /// <summary>
		 /// SupportByVersion Access 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 12,14,15,16)]
		 acPropertyCaption = 9,

		 /// <summary>
		 /// SupportByVersion Access 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 14,15,16)]
		 acPropertyValue = 10
	}
}