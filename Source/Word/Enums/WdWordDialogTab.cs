using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff844913.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdWordDialogTab
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabView = 204,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabGeneral = 203,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabEdit = 224,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabPrint = 208,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabSave = 209,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabProofread = 211,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>386</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabTrackChanges = 386,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabUserInfo = 213,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>525</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabCompatibility = 525,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>739</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabTypography = 739,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>225</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabFileLocations = 225,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>790</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabFuzzy = 790,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>786</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabHangulHanjaConversion = 786,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1029</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsOptionsTabBidi = 1029,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>150000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFilePageSetupTabMargins = 150000,

		 /// <summary>
		 /// SupportByVersion Word 9
		 /// </summary>
		 /// <remarks>150001</remarks>
		 [SupportByVersionAttribute("Word", 9)]
		 wdDialogFilePageSetupTabPaperSize = 150001,

		 /// <summary>
		 /// SupportByVersion Word 9
		 /// </summary>
		 /// <remarks>150002</remarks>
		 [SupportByVersionAttribute("Word", 9)]
		 wdDialogFilePageSetupTabPaperSource = 150002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>150003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFilePageSetupTabLayout = 150003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>150004</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFilePageSetupTabCharsLines = 150004,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>200000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertSymbolTabSymbols = 200000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>200001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertSymbolTabSpecialCharacters = 200001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>300000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogNoteOptionsTabAllFootnotes = 300000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>300001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogNoteOptionsTabAllEndnotes = 300001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertIndexAndTablesTabIndex = 400000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertIndexAndTablesTabTableOfContents = 400001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertIndexAndTablesTabTableOfFigures = 400002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>400003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogInsertIndexAndTablesTabTableOfAuthorities = 400003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogOrganizerTabStyles = 500000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogOrganizerTabAutoText = 500001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogOrganizerTabCommandBars = 500002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>500003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogOrganizerTabMacros = 500003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>600000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatFontTabFont = 600000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>600001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatFontTabCharacterSpacing = 600001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>600002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatFontTabAnimation = 600002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>700000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBordersAndShadingTabBorders = 700000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>700001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBordersAndShadingTabPageBorder = 700001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>700002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBordersAndShadingTabShading = 700002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>800000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsEnvelopesAndLabelsTabEnvelopes = 800000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>800001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsEnvelopesAndLabelsTabLabels = 800001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1000000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatParagraphTabIndentsAndSpacing = 1000000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1000001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatParagraphTabTextFlow = 1000001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1000002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatParagraphTabTeisai = 1000002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabColorsAndLines = 1200000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabSize = 1200001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabPosition = 1200002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabWrapping = 1200003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200004</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabPicture = 1200004,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200005</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabTextbox = 1200005,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200006</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabWeb = 1200006,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1200007</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatDrawingObjectTabHR = 1200007,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1400000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoCorrectExceptionsTabFirstLetter = 1400000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1400001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoCorrectExceptionsTabInitialCaps = 1400001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1400002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = 1400002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1400003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoCorrectExceptionsTabIac = 1400003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1500000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBulletsAndNumberingTabBulleted = 1500000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1500001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBulletsAndNumberingTabNumbered = 1500001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1500002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogFormatBulletsAndNumberingTabOutlineNumbered = 1500002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1600000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogLetterWizardTabLetterFormat = 1600000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1600001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogLetterWizardTabRecipientInfo = 1600001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1600002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogLetterWizardTabOtherElements = 1600002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1600003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogLetterWizardTabSenderInfo = 1600003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1700000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoManagerTabAutoCorrect = 1700000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1700001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoManagerTabAutoFormatAsYouType = 1700001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1700002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoManagerTabAutoText = 1700002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1700003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogToolsAutoManagerTabAutoFormat = 1700003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1900000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogEmailOptionsTabSignature = 1900000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1900001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogEmailOptionsTabStationary = 1900001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1900002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogEmailOptionsTabQuoting = 1900002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000000</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogWebOptionsGeneral = 2000000,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000001</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogWebOptionsFiles = 2000001,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000002</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogWebOptionsPictures = 2000002,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000003</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogWebOptionsEncoding = 2000003,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000004</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDialogWebOptionsFonts = 2000004,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogToolsOptionsTabSecurity = 1361,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>150001</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogFilePageSetupTabPaper = 150001,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1700004</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogToolsAutoManagerTabSmartTags = 1700004,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1800000</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogTablePropertiesTabTable = 1800000,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1800001</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogTablePropertiesTabRow = 1800001,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1800002</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogTablePropertiesTabColumn = 1800002,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1800003</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogTablePropertiesTabCell = 1800003,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2000000</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogWebOptionsBrowsers = 2000000,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1266</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDialogToolsOptionsTabAcetate = 1266,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2100000</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdDialogTemplates = 2100000,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2100001</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdDialogTemplatesXMLSchema = 2100001,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2100002</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdDialogTemplatesXMLExpansionPacks = 2100002,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2100003</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdDialogTemplatesLinkedCSS = 2100003,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>2200000</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDialogStyleManagementTabEdit = 2200000,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>2200001</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDialogStyleManagementTabRecommend = 2200001,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>2200002</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDialogStyleManagementTabRestrict = 2200002
	}
}