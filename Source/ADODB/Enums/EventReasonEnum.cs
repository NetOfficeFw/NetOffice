using System;
using NetOffice;
namespace NetOffice.ADODBApi.Enums
{
	 /// <summary>
	 /// SupportByVersion ADODB 2.1, 2.5
	 /// </summary>
	[SupportByVersionAttribute("ADODB", 2.1,2.5)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum EventReasonEnum
	{
		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnAddNew = 1,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnDelete = 2,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnUpdate = 3,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnUndoUpdate = 4,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnUndoAddNew = 5,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnUndoDelete = 6,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnRequery = 7,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnResynch = 8,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnClose = 9,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnMove = 10,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnFirstChange = 11,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnMoveFirst = 12,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnMoveNext = 13,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnMovePrevious = 14,

		 /// <summary>
		 /// SupportByVersion ADODB 2.1, 2.5
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("ADODB", 2.1,2.5)]
		 adRsnMoveLast = 15
	}
}