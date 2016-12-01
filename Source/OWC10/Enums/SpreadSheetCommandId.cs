using System;
using NetOffice;
namespace NetOffice.OWC10Api.Enums
{
	 /// <summary>
	 /// SupportByVersion OWC10 1
	 /// </summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum SpreadSheetCommandId
	{
		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1000</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandUndo = 1000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandCut = 1001,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1002</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandCopy = 1002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1003</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandPaste = 1003,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1004</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExport = 1004,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1005</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandProperties = 1005,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1006</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandHelp = 1006,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandAbout = 1007,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2000</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSortAsc = 2000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2030</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSortAscLast = 2030,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2031</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSortDesc = 2031,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>2061</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSortDescLast = 2061,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10000</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandAutosum = 10000,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10001</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandAutoFilter = 10001,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10002</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandClear = 10002,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1052</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandBold = 1052,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1053</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandItalic = 1053,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>1054</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandUnderline = 1054,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10006</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandDeleteRows = 10006,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10007</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandDeleteCols = 10007,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10008</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandInsertRows = 10008,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10009</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandInsertCols = 10009,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10010</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandRecalcForce = 10010,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10011</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSelectRow = 10011,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10012</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSelectCol = 10012,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10013</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSelectAll = 10013,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10014</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveLeft = 10014,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10015</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveUp = 10015,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10016</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveRight = 10016,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10017</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveDown = 10017,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10018</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandScrollLeft = 10018,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10019</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandScrollUp = 10019,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10020</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandScrollRight = 10020,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10021</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandScrollDown = 10021,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10022</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveNext = 10022,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10023</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMovePrevious = 10023,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10024</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandTabNext = 10024,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10025</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandTabPrevious = 10025,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10026</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToEndLeft = 10026,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10027</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToEndUp = 10027,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10028</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToEndRight = 10028,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10029</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToEndDown = 10029,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10030</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandLeft = 10030,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10031</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandUp = 10031,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10032</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandRight = 10032,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10033</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandDown = 10033,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10034</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToEndLeft = 10034,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10035</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToEndUp = 10035,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10036</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToEndRight = 10036,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10037</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToEndDown = 10037,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10038</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandEnterEditMode = 10038,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10039</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandShowContextMenu = 10039,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10040</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandToggleToolbar = 10040,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10041</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandEscape = 10041,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10042</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToLast = 10042,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10043</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToLast = 10043,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10044</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToLastInRow = 10044,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10045</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMovePageDown = 10045,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10046</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandPageDown = 10046,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10047</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMovePageUp = 10047,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10048</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandPageUp = 10048,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10062</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMovePageRight = 10062,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10063</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandPageRight = 10063,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10064</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMovePageLeft = 10064,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10065</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandPageLeft = 10065,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10049</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToOrigin = 10049,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10050</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToOrigin = 10050,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10051</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMoveToHome = 10051,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10052</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandToHome = 10052,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10053</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandExpandMenu = 10053,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10054</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandEat = 10054,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10055</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandNextSheet = 10055,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10056</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandPrevSheet = 10056,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10057</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandNewSheet = 10057,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10058</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSelectArray = 10058,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10067</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandSelectArraySilent = 10067,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10059</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandRecalc = 10059,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10060</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandRefresh = 10060,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10061</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandRefreshAll = 10061,

		 /// <summary>
		 /// SupportByVersion OWC10 1
		 /// </summary>
		 /// <remarks>10066</remarks>
		 [SupportByVersionAttribute("OWC10", 1)]
		 ssCommandMakeActiveCellVisible = 10066
	}
}