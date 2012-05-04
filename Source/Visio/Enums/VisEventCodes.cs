using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisEventCodes
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>-1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visScopeIDInval = -1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeInval = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocSave = 3,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocSaveAs = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocRunning = 5,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeDocDesign = 6,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefDocSave = 7,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefDocSaveAs = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelDocClose = 9,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelDocClose = 10,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>200</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefForcedFlush = 200,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>201</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeAfterForcedFlush = 201,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>202</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeEnterScope = 202,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeExitScope = 203,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelQuit = 204,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>205</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelQuit = 205,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelSuspend = 206,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>207</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelSuspend = 207,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBeforeSuspend = 208,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeAfterResume = 209,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelStyleDel = 300,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>301</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelStyleDel = 301,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>400</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelMasterDel = 400,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>401</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelMasterDel = 401,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>500</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelPageDel = 500,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>501</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelPageDel = 501,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>701</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeWinSelChange = 701,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>702</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefWinSelDel = 702,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>703</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefWinPageTurn = 703,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>704</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeWinPageTurn = 704,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>705</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeViewChanged = 705,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>706</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelWinClose = 706,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>707</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelWinClose = 707,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>708</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeWinOnAddonKeyMSG = 708,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>801</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeShapeDelete = 801,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>802</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeShapeParentChange = 802,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>803</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeShapeBeforeTextEdit = 803,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>804</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeShapeExitTextEdit = 804,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>901</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeBefSelDel = 901,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>902</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeSelAdded = 902,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>903</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelSelDel = 903,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>904</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelSelDel = 904,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>905</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelUngroup = 905,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>906</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelUngroup = 906,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>907</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeQueryCancelConvertToGroup = 907,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>908</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCancelConvertToGroup = 908,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>32768</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtAdd = 32768,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>16384</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtDel = 16384,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8192</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtMod = 8192,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtWindow = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtDoc = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtStyle = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtMaster = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtPage = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtLayer = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtShape = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtText = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtConnect = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtSection = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtRow = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2048</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCell = 2048,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtFormula = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4096</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtApp = 4096,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtAppActivate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtAppDeactivate = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtObjActivate = 4,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtObjDeactivate = 8,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtBeforeQuit = 16,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtBeforeModal = 32,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtAfterModal = 64,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtWinActivate = 128,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtMarker = 256,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>512</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtNonePending = 512,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1024</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtIdle = 1024,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>28672</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCode1stUser = 28672,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>32767</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeLastUser = 32767,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeCreate = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeOpen = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visActCodeRunAddon = 1,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visActCodeAdvise = 2,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtIdMostRecent = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>709</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeMouseDown = 709,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>710</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeMouseMove = 710,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>711</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeMouseUp = 711,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>712</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeKeyDown = 712,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>713</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeKeyPress = 713,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14
		 /// </summary>
		 /// <remarks>714</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14)]
		 visEvtCodeKeyUp = 714,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtDataRecordset = 32,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>805</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtShapeLinkAdded = 805,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>806</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtShapeLinkDeleted = 806,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>807</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtShapeDataGraphicChanged = 807,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>909</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeQueryCancelSelGroup = 909,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>910</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeCancelSelGroup = 910,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtRemoveHiddenInformation = 11,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>210</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeQueryCancelSuspendEvents = 210,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeCancelSuspendEvents = 211,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>212</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeBeforeSuspendEvents = 212,

		 /// <summary>
		 /// SupportByVersion Visio 12, 14
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersionAttribute("Visio", 12,14)]
		 visEvtCodeAfterResumeEvents = 213,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>502</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeContainerRelationshipAdded = 502,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>503</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeContainerRelationshipDeleted = 503,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>504</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeCalloutRelationshipAdded = 504,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeCalloutRelationshipDeleted = 505,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeSelectionMovedToSubprocess = 12,

		 /// <summary>
		 /// SupportByVersion Visio 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Visio", 14)]
		 visEvtCodeRuleSetValidated = 13
	}
}