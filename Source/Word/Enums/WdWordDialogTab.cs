﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.WdWordDialogTab"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum WdWordDialogTab
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabView = 204,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabGeneral = 203,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabEdit = 224,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabPrint = 208,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabSave = 209,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabProofread = 211,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>386</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabTrackChanges = 386,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabUserInfo = 213,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>525</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabCompatibility = 525,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>739</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabTypography = 739,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>225</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabFileLocations = 225,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>790</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabFuzzy = 790,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>786</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabHangulHanjaConversion = 786,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1029</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabBidi = 1029,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabMargins = 150000,

		 /// <summary>
		 /// SupportByVersion Word 9
		 /// </summary>
		 /// <remarks>150001</remarks>
		 [SupportByVersion("Word", 9)]
		 wdDialogFilePageSetupTabPaperSize = 150001,

		 /// <summary>
		 /// SupportByVersion Word 9
		 /// </summary>
		 /// <remarks>150002</remarks>
		 [SupportByVersion("Word", 9)]
		 wdDialogFilePageSetupTabPaperSource = 150002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabLayout = 150003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150004</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabCharsLines = 150004,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>200000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertSymbolTabSymbols = 200000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>200001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertSymbolTabSpecialCharacters = 200001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>300000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogNoteOptionsTabAllFootnotes = 300000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>300001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogNoteOptionsTabAllEndnotes = 300001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertIndexAndTablesTabIndex = 400000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertIndexAndTablesTabTableOfContents = 400001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertIndexAndTablesTabTableOfFigures = 400002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>400003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogInsertIndexAndTablesTabTableOfAuthorities = 400003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogOrganizerTabStyles = 500000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogOrganizerTabAutoText = 500001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogOrganizerTabCommandBars = 500002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>500003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogOrganizerTabMacros = 500003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>600000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatFontTabFont = 600000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>600001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatFontTabCharacterSpacing = 600001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>600002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatFontTabAnimation = 600002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>700000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBordersAndShadingTabBorders = 700000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>700001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBordersAndShadingTabPageBorder = 700001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>700002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBordersAndShadingTabShading = 700002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>800000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsEnvelopesAndLabelsTabEnvelopes = 800000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>800001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsEnvelopesAndLabelsTabLabels = 800001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatParagraphTabIndentsAndSpacing = 1000000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatParagraphTabTextFlow = 1000001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1000002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatParagraphTabTeisai = 1000002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabColorsAndLines = 1200000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabSize = 1200001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabPosition = 1200002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabWrapping = 1200003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200004</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabPicture = 1200004,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200005</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabTextbox = 1200005,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200006</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabWeb = 1200006,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1200007</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatDrawingObjectTabHR = 1200007,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1400000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoCorrectExceptionsTabFirstLetter = 1400000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1400001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoCorrectExceptionsTabInitialCaps = 1400001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1400002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = 1400002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1400003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoCorrectExceptionsTabIac = 1400003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1500000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBulletsAndNumberingTabBulleted = 1500000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1500001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBulletsAndNumberingTabNumbered = 1500001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1500002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogFormatBulletsAndNumberingTabOutlineNumbered = 1500002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1600000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogLetterWizardTabLetterFormat = 1600000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1600001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogLetterWizardTabRecipientInfo = 1600001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1600002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogLetterWizardTabOtherElements = 1600002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1600003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogLetterWizardTabSenderInfo = 1600003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoManagerTabAutoCorrect = 1700000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoManagerTabAutoFormatAsYouType = 1700001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoManagerTabAutoText = 1700002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogToolsAutoManagerTabAutoFormat = 1700003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1900000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogEmailOptionsTabSignature = 1900000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1900001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogEmailOptionsTabStationary = 1900001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1900002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogEmailOptionsTabQuoting = 1900002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000000</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogWebOptionsGeneral = 2000000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000001</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogWebOptionsFiles = 2000001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000002</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogWebOptionsPictures = 2000002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000003</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogWebOptionsEncoding = 2000003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000004</remarks>
		 [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		 wdDialogWebOptionsFonts = 2000004,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabSecurity = 1361,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150001</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogFilePageSetupTabPaper = 150001,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1700004</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogToolsAutoManagerTabSmartTags = 1700004,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1800000</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogTablePropertiesTabTable = 1800000,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1800001</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogTablePropertiesTabRow = 1800001,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1800002</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogTablePropertiesTabColumn = 1800002,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1800003</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogTablePropertiesTabCell = 1800003,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2000000</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogWebOptionsBrowsers = 2000000,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1266</remarks>
		 [SupportByVersion("Word", 10,11,12,14,15,16)]
		 wdDialogToolsOptionsTabAcetate = 1266,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2100000</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdDialogTemplates = 2100000,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2100001</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdDialogTemplatesXMLSchema = 2100001,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2100002</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdDialogTemplatesXMLExpansionPacks = 2100002,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2100003</remarks>
		 [SupportByVersion("Word", 11,12,14,15,16)]
		 wdDialogTemplatesLinkedCSS = 2100003,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2200000</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdDialogStyleManagementTabEdit = 2200000,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2200001</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdDialogStyleManagementTabRecommend = 2200001,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2200002</remarks>
		 [SupportByVersion("Word", 12,14,15,16)]
		 wdDialogStyleManagementTabRestrict = 2200002
	}
}