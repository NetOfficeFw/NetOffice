using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisEventCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visScopeIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeInval = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocSave = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocSaveAs = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocRunning = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeDocDesign = 6,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefDocSave = 7,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefDocSaveAs = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelDocClose = 9,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelDocClose = 10,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefForcedFlush = 200,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeAfterForcedFlush = 201,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeEnterScope = 202,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeExitScope = 203,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelQuit = 204,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelQuit = 205,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelSuspend = 206,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>207</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelSuspend = 207,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBeforeSuspend = 208,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeAfterResume = 209,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelStyleDel = 300,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>301</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelStyleDel = 301,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelMasterDel = 400,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>401</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelMasterDel = 401,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelPageDel = 500,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelPageDel = 501,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeWinSelChange = 701,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefWinSelDel = 702,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>703</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefWinPageTurn = 703,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>704</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeWinPageTurn = 704,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>705</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeViewChanged = 705,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>706</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelWinClose = 706,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>707</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelWinClose = 707,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>708</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeWinOnAddonKeyMSG = 708,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>801</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeShapeDelete = 801,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>802</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeShapeParentChange = 802,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>803</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeShapeBeforeTextEdit = 803,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>804</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeShapeExitTextEdit = 804,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>901</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeBefSelDel = 901,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>902</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeSelAdded = 902,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>903</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelSelDel = 903,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>904</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelSelDel = 904,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>905</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelUngroup = 905,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>906</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelUngroup = 906,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>907</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeQueryCancelConvertToGroup = 907,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>908</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCancelConvertToGroup = 908,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtAdd = 32768,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtDel = 16384,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtMod = 8192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtWindow = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtDoc = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtStyle = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtMaster = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtPage = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtLayer = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtShape = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtText = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtConnect = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtSection = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtRow = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCell = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtFormula = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtApp = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtAppActivate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtAppDeactivate = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtObjActivate = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtObjDeactivate = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtBeforeQuit = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtBeforeModal = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtAfterModal = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtWinActivate = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtMarker = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtNonePending = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtIdle = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28672</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCode1stUser = 28672,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32767</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeLastUser = 32767,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visActCodeRunAddon = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visActCodeAdvise = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtIdMostRecent = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>709</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeMouseDown = 709,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>710</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeMouseMove = 710,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>711</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeMouseUp = 711,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>712</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeKeyDown = 712,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>713</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeKeyPress = 713,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>714</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visEvtCodeKeyUp = 714,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtDataRecordset = 32,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>805</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtShapeLinkAdded = 805,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>806</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtShapeLinkDeleted = 806,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>807</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtShapeDataGraphicChanged = 807,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>909</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeQueryCancelSelGroup = 909,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>910</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeCancelSelGroup = 910,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtRemoveHiddenInformation = 11,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>210</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeQueryCancelSuspendEvents = 210,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeCancelSuspendEvents = 211,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>212</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeBeforeSuspendEvents = 212,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15)]
		 visEvtCodeAfterResumeEvents = 213,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>502</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeContainerRelationshipAdded = 502,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>503</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeContainerRelationshipDeleted = 503,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>504</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeCalloutRelationshipAdded = 504,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeCalloutRelationshipDeleted = 505,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeSelectionMovedToSubprocess = 12,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Visio", 14,15)]
		 visEvtCodeRuleSetValidated = 13,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>911</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visEvtCodeQueryCancelReplaceShapes = 911,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>912</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visEvtCodeCancelReplaceShapes = 912,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>913</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visEvtCodeBeforeReplaceShapes = 913,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>914</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visEvtCodeShapesReplaced = 914,

		 /// <summary>
		 /// SupportByVersion Visio 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Visio", 15)]
		 visEvtCodeAfterCoauthMerge = 14
	}
}