using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15, 16
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15,16)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisEventCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visScopeIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeInval = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocSave = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocSaveAs = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocRunning = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeDocDesign = 6,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefDocSave = 7,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefDocSaveAs = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelDocClose = 9,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelDocClose = 10,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefForcedFlush = 200,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeAfterForcedFlush = 201,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeEnterScope = 202,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeExitScope = 203,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelQuit = 204,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelQuit = 205,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelSuspend = 206,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>207</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelSuspend = 207,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBeforeSuspend = 208,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeAfterResume = 209,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelStyleDel = 300,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>301</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelStyleDel = 301,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelMasterDel = 400,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>401</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelMasterDel = 401,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelPageDel = 500,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelPageDel = 501,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeWinSelChange = 701,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefWinSelDel = 702,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>703</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefWinPageTurn = 703,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>704</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeWinPageTurn = 704,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>705</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeViewChanged = 705,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>706</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelWinClose = 706,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>707</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelWinClose = 707,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>708</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeWinOnAddonKeyMSG = 708,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>801</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeShapeDelete = 801,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>802</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeShapeParentChange = 802,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>803</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeShapeBeforeTextEdit = 803,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>804</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeShapeExitTextEdit = 804,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>901</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeBefSelDel = 901,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>902</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeSelAdded = 902,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>903</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelSelDel = 903,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>904</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelSelDel = 904,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>905</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelUngroup = 905,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>906</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelUngroup = 906,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>907</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeQueryCancelConvertToGroup = 907,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>908</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCancelConvertToGroup = 908,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtAdd = 32768,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtDel = 16384,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtMod = 8192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtWindow = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtDoc = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtStyle = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtMaster = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtPage = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtLayer = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtShape = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtText = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtConnect = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtSection = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtRow = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCell = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtFormula = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtApp = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtAppActivate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtAppDeactivate = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtObjActivate = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtObjDeactivate = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtBeforeQuit = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtBeforeModal = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtAfterModal = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtWinActivate = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtMarker = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtNonePending = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtIdle = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28672</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCode1stUser = 28672,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32767</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeLastUser = 32767,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visActCodeRunAddon = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visActCodeAdvise = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtIdMostRecent = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>709</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeMouseDown = 709,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>710</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeMouseMove = 710,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>711</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeMouseUp = 711,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>712</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeKeyDown = 712,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>713</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeKeyPress = 713,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>714</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15,16)]
		 visEvtCodeKeyUp = 714,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtDataRecordset = 32,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>805</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtShapeLinkAdded = 805,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>806</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtShapeLinkDeleted = 806,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>807</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtShapeDataGraphicChanged = 807,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>909</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeQueryCancelSelGroup = 909,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>910</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeCancelSelGroup = 910,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtRemoveHiddenInformation = 11,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>210</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeQueryCancelSuspendEvents = 210,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeCancelSuspendEvents = 211,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>212</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeBeforeSuspendEvents = 212,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersionAttribute("Visio", 12,14,15,16)]
		 visEvtCodeAfterResumeEvents = 213,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>502</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeContainerRelationshipAdded = 502,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>503</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeContainerRelationshipDeleted = 503,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>504</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeCalloutRelationshipAdded = 504,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeCalloutRelationshipDeleted = 505,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeSelectionMovedToSubprocess = 12,

		 /// <summary>
		 /// SupportByVersion Visio 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Visio", 14,15,16)]
		 visEvtCodeRuleSetValidated = 13,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>911</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visEvtCodeQueryCancelReplaceShapes = 911,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>912</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visEvtCodeCancelReplaceShapes = 912,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>913</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visEvtCodeBeforeReplaceShapes = 913,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>914</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visEvtCodeShapesReplaced = 914,

		 /// <summary>
		 /// SupportByVersion Visio 15,16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Visio", 15, 16)]
		 visEvtCodeAfterCoauthMerge = 14
	}
}