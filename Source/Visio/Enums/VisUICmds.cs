using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisUICmds
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFirst = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65535</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLast = 65535,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHierarchical = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1001</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNew = 1001,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1002</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileOpen = 1002,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1003</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileClose = 1003,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1004</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileSave = 1004,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1005</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileSaveAs = 1005,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1006</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileSaveWorkspace = 1006,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileImport = 1007,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1009</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileSummaryInfoDlg = 1009,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1010</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFilePrint = 1010,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1012</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile1 = 1012,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1013</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile2 = 1013,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1014</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile3 = 1014,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1015</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile4 = 1015,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1016</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileExit = 1016,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1017</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditUndo = 1017,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1018</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditRedo = 1018,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1019</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditRepeat = 1019,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1020</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditCut = 1020,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1021</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditCopy = 1021,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1022</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditPaste = 1022,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1023</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditClear = 1023,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditDuplicate = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1025</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUFEditSelectAll = 1025,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1026</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditSelectSpecial = 1026,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1027</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditPasteSpecial = 1027,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1028</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditPasteLink = 1028,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1029</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditOpenObject = 1029,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1030</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditLinks = 1030,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1031</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditInsertObject = 1031,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1032</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditInsertField = 1032,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1033</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewFitInWindow = 1033,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1034</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView75 = 1034,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1035</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView100 = 1035,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1036</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView150 = 1036,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1037</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView200 = 1037,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1038</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewCustom = 1038,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1039</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewRulers = 1039,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1040</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewGrid = 1040,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1041</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewGuides = 1041,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1042</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewConnections = 1042,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1043</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditFind = 1043,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1044</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewStatusBar = 1044,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1045</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectBringForward = 1045,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1046</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectBringToFront = 1046,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1047</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectSendBackward = 1047,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1048</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectSendToBack = 1048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1049</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectAlignObjects = 1049,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1050</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectConnectObjects = 1050,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1051</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectGroup = 1051,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1052</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectUngroup = 1052,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1053</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectAddToGroup = 1053,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1054</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectRemoveFromGroup = 1054,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1055</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectConvertToGroup = 1055,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1056</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectRotate90 = 1056,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1057</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectFlipVertical = 1057,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1058</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectFlipHorizontal = 1058,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1059</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectReverse = 1059,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1060</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectUnion = 1060,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1061</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectCombine = 1061,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1062</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectFragment = 1062,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1063</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatStyle = 1063,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1064</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatDefineStyles = 1064,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1065</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatLine = 1065,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1066</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatFill = 1066,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1067</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatText = 1067,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1068</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatParagraph = 1068,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1069</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatTabs = 1069,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1070</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatBlock = 1070,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1071</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatBehavior = 1071,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1072</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatProtection = 1072,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1073</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatSpecial = 1073,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1074</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsEditDrawing = 1074,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1075</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsEditBackground = 1075,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1076</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsPageSetup = 1076,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1077</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsGoToDrawing = 1077,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1078</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsNewPage = 1078,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1079</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsDeletePages = 1079,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1080</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsReorderPages = 1080,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1081</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsPreferences = 1081,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1082</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsColorPaletteDlg = 1082,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1083</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsProtectDocument = 1083,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1084</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOptionsSnapGlueSetup = 1084,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1085</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowNewWindow = 1085,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1086</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowCascadeAll = 1086,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1087</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowTileAll = 1087,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1088</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowShowShapeSheet = 1088,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1089</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowShowMasterObjects = 1089,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1090</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRunAddOnMenu = 1090,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1091</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdWindowShowDrawPage = 1091,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1092</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpContents = 1092,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1093</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDecreaseIndent = 1093,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1094</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIncreaseIndent = 1094,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1095</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDecreaseParaSpacing = 1095,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIncreaseParaSpacing = 1096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1097</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpStencil = 1097,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1098</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextRotate90 = 1098,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1099</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpQuickTour = 1099,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1100</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpAboutVisio = 1100,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1101</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenEditIcon = 1101,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1102</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenEditDrawing = 1102,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1103</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenNameMaster = 1103,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1104</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenNewMaster = 1104,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1105</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenImageMaster = 1105,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1106</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenCleanup = 1106,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1107</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWShowValues = 1107,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1108</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWShowFormulas = 1108,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1109</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWShowSectionsDlg = 1109,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1110</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWPasteNameDlg = 1110,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1111</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWPasteFunctionDlg = 1111,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1112</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWInsertRow = 1112,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1113</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWInsertRowAfter = 1113,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1114</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWChangeRowTypeDlg = 1114,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1115</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWDeleteRow = 1115,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1116</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWAddSectionDlg = 1116,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1117</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWDeleteSection = 1117,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1118</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatDoubleClick = 1118,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1121</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawTextStyle = 1121,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1122</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawLineStyle = 1122,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1123</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawFillStyle = 1123,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1124</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawSnap = 1124,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1125</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawGlue = 1125,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1126</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawZoom = 1126,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextStyle = 1128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1129</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextFont = 1129,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1130</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextSize = 1130,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1131</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextBold = 1131,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1132</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextItalic = 1132,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1133</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextSmallCaps = 1133,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1134</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextSuperscript = 1134,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1135</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextSubscript = 1135,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1136</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextUline = 1136,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1139</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWCancel = 1139,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1140</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWAccept = 1140,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1141</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWFormula = 1141,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1142</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWShowToggle = 1142,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1143</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconLeftColor = 1143,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1144</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconRightColor = 1144,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1145</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconPencilTool = 1145,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1146</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRecalcObjectWH = 1146,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1147</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTurnToPrevPage = 1147,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1148</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTurnToNextPage = 1148,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1179</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditReplace = 1179,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1180</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawAddGuide = 1180,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1181</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAddTextShape = 1181,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1182</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawRect = 1182,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1183</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawOval = 1183,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1184</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDragDuplicate = 1184,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1185</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMoveObject = 1185,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1186</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMove1D = 1186,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1187</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMove2D = 1187,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1188</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSize1D = 1188,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1189</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSize2D = 1189,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1190</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRotateObject = 1190,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1192</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCropObject = 1192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1193</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPanObject = 1193,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1194</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSizeTextBlock = 1194,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1196</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectLeft = 1196,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1197</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectCenter = 1197,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1198</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectRight = 1198,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1199</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectTop = 1199,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectMiddle = 1200,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1201</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignObjectBottom = 1201,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1202</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCenterDrawing = 1202,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1213</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDeselectAll = 1213,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1214</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextEditState = 1214,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1215</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomPt = 1215,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1216</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomIn = 1216,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1217</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomOut = 1217,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1218</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomArea = 1218,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1219</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRPointerTool = 1219,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1220</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRPencilTool = 1220,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1221</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRLineTool = 1221,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1222</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRQtrArcTool = 1222,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1223</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRRectTool = 1223,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1224</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDROvalTool = 1224,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1225</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRConnectorTool = 1225,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1226</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRConnectionTool = 1226,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1227</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRTextTool = 1227,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1228</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRRotateTool = 1228,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1230</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectDistributeDlg = 1230,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1231</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeHSpace = 1231,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1232</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeLeft = 1232,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1233</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeCenter = 1233,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1234</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeRight = 1234,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1235</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeVSpace = 1235,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1236</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeTop = 1236,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1237</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeMiddle = 1237,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1238</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDistributeBottom = 1238,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1241</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdUpdateContentCache = 1241,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1243</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDropOnText = 1243,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1244</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDropOnStencil = 1244,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1246</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDropOnPage = 1246,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1250</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowCollapse = 1250,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1251</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowExpand = 1251,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1252</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowSelect = 1252,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1253</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowDeselect = 1253,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1263</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAddConnectPt = 1263,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1264</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdModConnectPt = 1264,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1265</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDelConnectPt = 1265,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1266</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAddControlPt = 1266,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1267</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdModControlPt = 1267,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1268</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDelControlPt = 1268,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1269</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMovConnectPt = 1269,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1270</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsSpelling = 1270,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1271</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatPainter = 1271,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1274</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPageMeasureUnitsDlg = 1274,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1279</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView50 = 1279,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1280</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdView400 = 1280,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1282</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertDataMap = 1282,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1292</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSendAsMail = 1292,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1309</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeActions = 1309,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1311</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDRSplineTool = 1311,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1312</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatCustPropEdit = 1312,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1318</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRulerGridDlg = 1318,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1333</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatShadow = 1333,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1334</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatCorners = 1334,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1335</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsInventory = 1335,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1343</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMasterSetup = 1343,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1354</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsArrayShapesAddOn = 1354,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1355</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLineWeight = 1355,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1356</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLinePattern = 1356,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1357</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLineEnds = 1357,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1358</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLineCornerStyle = 1358,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1359</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLineColor = 1359,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCloseWindow = 1361,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1379</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetFillShadow = 1379,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1380</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowShowSection = 1380,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1381</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowPasteName = 1381,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1382</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowPasteFunction = 1382,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1383</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowChangeRowType = 1383,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1384</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowAddSection = 1384,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1385</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetFillColor = 1385,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1386</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpMode = 1386,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1387</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOffsetDlg = 1387,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1388</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDesignMode = 1388,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1389</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeExplorer = 1389,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1399</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetFillPattern = 1399,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1404</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetCharColor = 1404,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1405</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetCharSizeUp = 1405,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1406</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetCharSizeDown = 1406,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1407</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextHAlignLeft = 1407,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1408</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextHAlignCenter = 1408,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1409</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextHAlignRight = 1409,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1412</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextHAlignJustify = 1412,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1413</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextVAlignTop = 1413,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1414</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextVAlignMiddle = 1414,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1422</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextVAlignBottom = 1422,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1424</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStampTool = 1424,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1425</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectInfoDlg = 1425,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1428</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectHelp = 1428,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1439</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditConvertObject = 1439,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1442</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileOpenStencil = 1442,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1443</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPrintPage = 1443,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1444</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWShapeActionDlg = 1444,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1446</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLayerDlg = 1446,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1448</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLayerSetupDlg = 1448,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1449</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCropTool = 1449,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1451</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextBlockTool = 1451,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1452</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenClose = 1452,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1453</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIntersect = 1453,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1454</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSubtract = 1454,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1458</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenActivate = 1458,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1480</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenIconAndName = 1480,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1481</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenIconOnly = 1481,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1482</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenNameOnly = 1482,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1483</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenAutoArrange = 1483,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1484</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRunAddOnDlg = 1484,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1490</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPrintPreview = 1490,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1491</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOpenInVisio = 1491,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1492</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFullScreenMode = 1492,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1493</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLayoutDynamic = 1493,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1494</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRotate90Clockwise = 1494,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1495</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomLast = 1495,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1496</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomPageWidth = 1496,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1497</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertClipArt = 1497,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1498</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertWordArt = 1498,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1499</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertMicrosoftGraph = 1499,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1500</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolbarsDlg = 1500,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1501</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertComment = 1501,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1502</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMoveComment = 1502,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1503</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOpenCommentForEdit = 1503,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1504</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMSOInsertSymbol = 1504,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1505</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMSOInsertSymbolDlg = 1505,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1506</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETAddToFavorites = 1506,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1509</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdViewPageBreaks = 1509,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomSingleTile = 1512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1513</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPreviousTile = 1513,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1514</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNextTile = 1514,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1515</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFirstTile = 1515,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1516</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLastTile = 1516,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1521</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertAutoCADAddOn = 1521,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1522</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertControlDlg = 1522,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1533</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdJoin = 1533,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1534</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTrim = 1534,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1536</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDlgCustomFit = 1536,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1538</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFitCurve = 1538,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1543</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconBucketTool = 1543,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1544</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconLassoTool = 1544,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1545</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdIconSelectNet = 1545,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1561</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile5 = 1561,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1569</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile6 = 1569,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1570</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile7 = 1570,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1571</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile8 = 1571,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1572</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileLastFile9 = 1572,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1574</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsLayoutShapesDlg = 1574,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1576</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsRunVBE = 1576,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1577</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsMacroDlg = 1577,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1579</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankDrawing = 1579,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1580</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewStencilDlg = 1580,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1582</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankStencil = 1582,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1583</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileChooseTemplates = 1583,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1584</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdProgRefHelp = 1584,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1585</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInsertHyperLink = 1585,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1586</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpTemplates = 1586,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1588</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEmailRouting = 1588,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1589</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSendToExchange = 1589,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1590</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRemoveVBAFromActiveDoc = 1590,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1595</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETUserSearchPage = 1595,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1596</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETVisioHomePage = 1596,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1598</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETGoForward = 1598,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1599</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETGoBack = 1599,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1601</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdOpenActiveObject = 1601,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1602</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCancelInPlaceEditing = 1602,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1604</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETVisioSolutionsLibrary = 1604,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1605</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETKnowledgeBase = 1605,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1606</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETDiagrammingResources = 1606,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1607</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETOpenHlink = 1607,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1608</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETOpenHlinkNewWnd = 1608,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1609</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETDeleteHlink = 1609,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1610</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETCopyHyperlink = 1610,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1611</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHyperlinkHier = 1611,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1619</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETEditHyperlink = 1619,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1620</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETPasteAsHyperlink = 1620,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1633</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdBullets = 1633,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1634</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeLayerToolbar = 1634,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1635</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdGoToPageToolbar = 1635,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1642</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatAllTextProps = 1642,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1645</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdBrowseSampleDrawings = 1645,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1646</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMSOInsertEquation = 1646,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1650</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdABarHide = 1650,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1651</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdABarToggleFloat = 1651,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1652</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdABarAutohide = 1652,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1653</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPanZoom = 1653,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1654</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPagesList = 1654,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1658</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCustProp = 1658,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1661</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkTool = 1661,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1664</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkCustomizePen = 1664,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1669</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapesWindow = 1669,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1670</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSizePos = 1670,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1671</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankDrawingMetric = 1671,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1672</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankDrawingUS = 1672,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1673</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankStencilMetric = 1673,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1674</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileNewBlankStencilUS = 1674,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1675</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCustomPropertySets = 1675,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1676</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenSave = 1676,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1677</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenSaveAs = 1677,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1678</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenProperties = 1678,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1679</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenEditToggle = 1679,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1680</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenEditOn = 1680,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1681</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenEditOff = 1681,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1682</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditUndoMultiple = 1682,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1683</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdEditRedoMultiple = 1683,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1684</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdABarAutoHeight = 1684,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1685</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeCommentDlg = 1685,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1686</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeComment = 1686,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1687</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatCustPropDef = 1687,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1688</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeCommentDelete = 1688,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1689</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHideDocumentStencil = 1689,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1690</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToggleDocumentStencil = 1690,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1695</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCustPropDefine = 1695,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1696</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnRerouteFreely = 1696,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1697</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnRerouteAsNeeded = 1697,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1698</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnRerouteNever = 1698,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1699</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPagePlow = 1699,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnRoutingStyle = 1700,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1702</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPlaceableShapeBehavior = 1702,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1703</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPageLineJumpCode_Disp = 1703,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1704</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPageLineJumpCode_None = 1704,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1705</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPageLineJumpCode_Horz = 1705,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1706</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPageLineJumpCode_Vert = 1706,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1707</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetPageLineJumpCode_Last = 1707,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1708</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_Page = 1708,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1709</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_Arc = 1709,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1710</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_Gap = 1710,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1711</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_Square = 1711,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1712</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_Triangle = 1712,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1713</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_2pt = 1713,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1714</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_3pt = 1714,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1715</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_4pt = 1715,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1716</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_5pt = 1716,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1717</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnLineJumpStyle_6pt = 1717,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1718</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSWExpandRow = 1718,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1719</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHyperlinkList = 1719,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1720</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHeaderFooter = 1720,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1721</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawingExplorer = 1721,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1726</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHideAllToolbars = 1726,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1741</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextStrikethrough = 1741,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1742</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawRegion = 1742,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1744</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetAddMarkup = 1744,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1765</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDynamicGrid = 1765,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1766</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRulSub = 1766,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1767</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdGrid = 1767,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1768</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAlignBox = 1768,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1769</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeGeo = 1769,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1771</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdGuides = 1771,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1772</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeHand = 1772,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1773</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeVert = 1773,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1774</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdConnPoints = 1774,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1775</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRecordNewMacro = 1775,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1776</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStartRecordingMacro = 1776,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1777</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStopRecordingMacro = 1777,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1778</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPauseRecordingMacro = 1778,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1779</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdResumeRecordingMacro = 1779,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1781</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSSWindowShowTraceWindow = 1781,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1785</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileSaveAsWebPage = 1785,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1787</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileCheckIn = 1787,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1788</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFileCheckOut = 1788,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1790</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPasteShortcut = 1790,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1791</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCreateShortcut = 1791,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1795</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdReOrderPage = 1795,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1796</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenDrawingExplorer = 1796,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1802</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETOfficeOnTheWeb = 1802,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1807</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolSnapLines = 1807,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1809</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpSearch = 1809,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1810</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextEditRuler = 1810,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1812</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCreateNewDrawing = 1812,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1822</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdHelpShapeBasics = 1822,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1829</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDynConnReroute = 1829,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1830</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeIntersect = 1830,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1831</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdINETVisioOnTheWeb = 1831,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1836</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdReviewerVisibilityAll = 1836,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1837</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnRerouteOnCrossover = 1837,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1857</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSaveForAutoRecover = 1857,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1858</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetHeaderFooter = 1858,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1859</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoClipboard = 1859,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1860</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoSearch = 1860,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1862</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextAllCaps = 1862,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1863</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextDoubleUline = 1863,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1864</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAppMaximize = 1864,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1865</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAppMinimize = 1865,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1866</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoAutoCorrectDlg = 1866,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1867</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeGalleryAddOn = 1867,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1868</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAcquireImages = 1868,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1869</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDropManyOnPage = 1869,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1870</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdObjectSwapEnds = 1870,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1871</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetIndexInStencil = 1871,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1872</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoAutoCorrect = 1872,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1873</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoAutoFormat = 1873,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1874</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeTransparencyDlg = 1874,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1875</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeTransparency = 1875,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1876</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsShowAddins = 1876,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1877</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdLicenseVerification = 1877,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1878</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRightDragMove = 1878,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1879</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRightDragCopy = 1879,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1880</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRightDragLink = 1880,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1881</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdRightDragCancel = 1881,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1882</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNMMeetNow = 1882,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1883</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdImagePropertiesDlg = 1883,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1884</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdToolsSecurity = 1884,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1885</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoMediaGallery = 1885,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1886</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNextWindow = 1886,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1887</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdImageProperties = 1887,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1888</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetLanguageDlg = 1888,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1889</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSpellingChange = 1889,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1890</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDetectAndRepair = 1890,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1891</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdExportDatabaseAddOn = 1891,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1892</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdStenIconAndDetail = 1892,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1893</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnAppearanceDefault = 1893,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1894</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnAppearanceStraight = 1894,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1895</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSetDynConnAppearanceCurved = 1895,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1896</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTaskPaneToggle = 1896,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1897</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNewFromExisting = 1897,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1898</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMsoCustomItem = 1898,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1899</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCreateEditMaster = 1899,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1900</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdBreakOLELink = 1900,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1901</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMDIMaximize = 1901,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1902</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMDIMinimize = 1902,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1903</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMDIRestore = 1903,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1904</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdAppRestore = 1904,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1905</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDeleteBackWord = 1905,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1906</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNewDefDocBlankDrawing = 1906,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1907</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSelectionModeRect = 1907,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1908</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSelectionModeLasso = 1908,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1909</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSelectionModeExtend = 1909,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1914</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdNextMarkup = 1914,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1915</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPreviousMarkup = 1915,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1916</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdMasterExplorer = 1916,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1917</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomInIgnoreSel = 1917,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1918</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdZoomOutIgnoreSel = 1918,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1919</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdReviewerVisibilityNone = 1919,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1920</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDeleteComment = 1920,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1939</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdReviewerPaneToggle = 1939,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1943</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdConnectorEffectRightAngle = 1943,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1944</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdConnectorEffectStraight = 1944,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1945</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdConnectorEffectCurved = 1945,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1946</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDrawingTools = 1946,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1951</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextDoubleStrikethrough = 1951,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1952</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdTextHAlignDistribute = 1952,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1955</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdFormatInkDlg = 1955,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1962</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdCheckForUpdates = 1962,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1963</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdPrivacySettings = 1963,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1964</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdContactUs = 1964,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1967</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdResearchLookUp = 1967,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1968</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdResearchTranslate = 1968,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1969</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdResearchPaneToggle = 1969,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1970</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkEraser = 1970,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1971</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkReviewPen = 1971,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1972</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdSharedWorkspacePaneToggle = 1972,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1973</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkStockPen0 = 1973,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1974</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkStockPen1 = 1974,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1975</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkStockPen2 = 1975,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1976</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkStockPen3 = 1976,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1977</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdInkStockPen4 = 1977,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1982</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdDiagramGallery = 1982,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1985</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visCmdShapeStudioAddon = 1985,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1925</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdSizeObjects = 1925,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1997</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdLinkRowToShape = 1997,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1998</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdAddDataRecordset = 1998,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1999</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDeleteDataRecordset = 1999,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2005</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdStenNamesUnderIcons = 2005,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2006</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdInsertTextBox = 2006,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2007</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdInsertVertTextBox = 2007,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2009</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdRHI = 2009,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2010</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdRHIDlg = 2010,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2011</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataSelectorDlg = 2011,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2012</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdViewDirectionToggle = 2012,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2013</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdViewLeftToRight = 2013,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2014</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdViewRightToLeft = 2014,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2017</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdApplyDataGraphic = 2017,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2019</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefreshDlg = 2019,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2021</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefresh = 2021,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2022</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefreshConfigDlg = 2022,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2024</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneDataGraphic = 2024,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2037</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRecordsetSetCommand = 2037,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2038</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRecordsetSetPrimaryKey = 2038,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2042</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdSpellingOptionsDlg = 2042,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2043</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataColumnSettingsDlg = 2043,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2044</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataExplorerWindow = 2044,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2045</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataAutoLinkWiz = 2045,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2046</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataAutoLink = 2046,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2047</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdApplyThemeToPage = 2047,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdApplyThemeToDoc = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2049</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdEditTheme = 2049,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2050</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDuplicateTheme = 2050,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2052</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDeleteTheme = 2052,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2053</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskTogglePreviewSize = 2053,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2054</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneThemeColors = 2054,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2055</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneThemeEffects = 2055,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2056</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdAllowThemes = 2056,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2057</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataUnlinkShape = 2057,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2058</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataUnlinkRow = 2058,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2061</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdUpdateColumnsInLinkedShapes = 2061,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2064</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdNewThemeEffects = 2064,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2065</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdNewThemeColors = 2065,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2067</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDeleteDataGraphic = 2067,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2068</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdRelayoutShapes = 2068,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2072</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRecordsetProperties = 2072,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2091</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdAutoConnectToggle = 2091,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2092</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdApplyDataGraphicAfterLink = 2092,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2094</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefreshAddConflict = 2094,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2095</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefreshDeleteConflict = 2095,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2098</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataAutoConnect = 2098,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2103</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDataRefreshResolveConflict = 2103,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2104</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTrustCenterDlg = 2104,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2105</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdAutoGenerateDataGraphics = 2105,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2106</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDuplicateDataGraphic = 2106,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2107</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdRemoveDataGraphicFromSel = 2107,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2108</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDropManyLinked = 2108,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2109</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdFileUndoCheckout = 2109,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2114</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdDeleteForwardWord = 2114,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2117</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdSaveAsFixedFormatDlg = 2117,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2119</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdRemoveThemeFromSel = 2119,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1896</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPane = 1896,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1939</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneReviewer = 1939,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1969</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneResearch = 1969,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1972</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdTaskPaneDocumentManagement = 1972,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1890</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visCmdOfficeDiagnostics = 1890,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2127</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile10 = 2127,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2128</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile11 = 2128,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2129</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile12 = 2129,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2130</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile13 = 2130,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2131</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile14 = 2131,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2132</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile15 = 2132,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2133</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile16 = 2133,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2134</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile17 = 2134,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2135</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile18 = 2135,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2136</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile19 = 2136,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2137</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileLastFile20 = 2137,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2141</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdOfficeCenterOptions = 2141,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2144</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertLabelControl = 2144,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2145</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInserTextBoxControl = 2145,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2146</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertSpinControl = 2146,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2147</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertPushButtonControl = 2147,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2148</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertImageControl = 2148,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2149</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertScrollBarControl = 2149,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2150</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertCheckBoxControl = 2150,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2151</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertRadioButtonControl = 2151,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2152</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertComboBoxControl = 2152,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2153</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertListBoxControl = 2153,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2154</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertToggleButtonControl = 2154,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2165</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertNewBackgroundPage = 2165,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2167</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShowShapeSheetShape = 2167,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2168</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShowShapeSheetPage = 2168,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2169</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShowShapeSheetDocument = 2169,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2170</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSetPageOrientation = 2170,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2171</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSetPageSize = 2171,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2172</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDropAndContain = 2172,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2173</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAddMemberToContainer = 2173,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2174</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertMemberIntoList = 2174,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2175</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdRemoveMemberFromContainer = 2175,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2176</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPageSizeDlg = 2176,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2178</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdResearchThesaurus = 2178,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2179</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPreviousCommentMarkup = 2179,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2180</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdNextCommentMarkup = 2180,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2181</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSetContainerProperties = 2181,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2188</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyThemeColors = 2188,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2189</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyThemeEffects = 2189,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2190</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdEditThemeColors = 2190,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2191</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdEditThemeEffects = 2191,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2195</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFitContainerToContents = 2195,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2196</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDropAndInsertIntoList = 2196,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2197</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdReorderList = 2197,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2199</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDeleteConnectors = 2199,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2201</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdMultipleFileImport = 2201,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2202</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertPageTab = 2202,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2204</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDisbandContainer = 2204,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2205</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFormatPictureAutobalance = 2205,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2212</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFormatPictureCompressionDlg = 2212,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2213</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdCloseInkToolsRibbonTab = 2213,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2219</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSelectContainerMembers = 2219,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2220</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdLockContainer = 2220,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2221</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPasteToLocation = 2221,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2222</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAutoAlignAndSpace = 2222,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2223</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAutoAlign = 2223,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2224</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAutoSpace = 2224,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2225</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdMoveOffPageBreaks = 2225,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2226</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDiagramRotateRight = 2226,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2227</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDiagramRotateLeft = 2227,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2228</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDiagramFlipVertical = 2228,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2229</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDiagramFlipHorizontal = 2229,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2231</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShowLineJumpsToggle = 2231,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2232</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdMinimizeRibbonToggle = 2232,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2233</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdLayoutSpacingDlg = 2233,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2238</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdCOMAddinsDlg = 2238,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2245</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdNewSubProcess = 2245,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2249</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdCreateSubProcessFromSel = 2249,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2251</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAddToAllContainers = 2251,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2252</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdRemoveFromAllContainers = 2252,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2253</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidateDiagram = 2253,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2254</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdIgnoreValidationIssue = 2254,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2255</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdReinstateValidationIssue = 2255,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2256</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdIgnoreValidationRule = 2256,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2257</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdStopIgnoringValidationRule = 2257,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2258</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShowIgnoredIssuesToggle = 2258,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2263</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesWindowToggle = 2263,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2265</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSetDiagramServices = 2265,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2266</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSetAutoSize = 2266,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2267</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAutoSizeDrawing = 2267,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2268</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAddToNewContainer = 2268,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2269</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdCollapseShapesWindow = 2269,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2270</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdListInsertBefore = 2270,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2271</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdListInsertAfter = 2271,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2278</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesArrangeByRule = 2278,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2279</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesArrangeByCategory = 2279,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2280</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesArrangeByPage = 2280,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2281</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesArrangeByIgnored = 2281,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2282</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssuesArrangeOriginalOrder = 2282,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2285</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyMainTheme = 2285,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2286</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdDropCallout = 2286,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2287</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdAssociateCallout = 2287,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2289</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyMainThemeToPage = 2289,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2290</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdActivateQuickShapesWnd = 2290,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2291</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdHideMoreShapes = 2291,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2293</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPublishToVisioServices = 2293,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2294</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPublishToProcessRepository = 2294,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2295</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdEditRedoOrRepeat = 2295,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2296</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyMainThemeToDocument = 2296,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2297</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdApplyThemeToNewShapesToggle = 2297,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2298</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsDrawingPreviousFileFormat = 2298,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2299</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsTemplate = 2299,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2300</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsPNG = 2300,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2301</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsJPG = 2301,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2302</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsEMF = 2302,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2303</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsSVG = 2303,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2304</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsVDX = 2304,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2305</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsDWG = 2305,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2306</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdFileSaveAsDrawing = 2306,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2326</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdValidationIssueNavigateToShape = 2326,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2331</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertLegendHorizontal = 2331,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2332</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPageSizeToFitDrawing = 2332,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2333</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPageAutoSizeToggle = 2333,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2335</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertLegendVertical1 = 2335,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2337</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdPostDrag = 2337,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2340</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdSpaceShapesAvoidPageBreaksToggle = 2340,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2344</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdShapeSearchWindowToggle = 2344,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2345</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdInsertClipArtDlg = 2345,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2346</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdRemoveMemberFromList = 2346,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2352</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdTranslateOptions = 2352,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2347</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdContainerAutoResizeOff = 2347,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2348</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdContainerAutoResizeExpandOnly = 2348,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2349</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdContainerAutoResizeExpandContract = 2349,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2361</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdNewForegroundPage = 2361,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>2363</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visCmdLanguagePreferencesDlg = 2363,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2051</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdReplaceShape = 2051,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2182</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdCheckCompatibility = 2182,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2311</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdFileSaveAsMacroDrawing = 2311,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2382</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdSetThemeBehavior = 2382,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2383</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdDuplicatePage = 2383,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2384</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdUpgradeThemeModel = 2384,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>2385</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visCmdCoauthMerging = 2385
	}
}