using System;
using LateBindingApi.Core;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdWordDialog
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogHelpAbout = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogHelpWordPerfectHelp = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>511</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogHelpWordPerfectHelpOptions = 511,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>322</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatChangeCase = 322,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>790</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsFuzzy = 790,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>228</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsWordCount = 228,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogDocumentStatistics = 78,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileNew = 79,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileOpen = 80,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeOpenDataSource = 81,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeOpenHeaderSource = 82,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>779</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeUseAddressBook = 779,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileSaveAs = 84,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileSummaryInfo = 86,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsTemplates = 87,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>222</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogOrganizer = 222,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFilePrint = 88,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>676</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMerge = 676,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>677</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeCheck = 677,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>681</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeQueryOptions = 681,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>569</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeFindRecord = 569,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4049</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertIf = 4049,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4053</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertNextIf = 4053,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4055</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertSkipIf = 4055,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4048</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertFillIn = 4048,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4047</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertAsk = 4047,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4054</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeInsertSet = 4054,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>680</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeHelper = 680,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>821</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogLetterWizard = 821,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFilePrintSetup = 97,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileFind = 99,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>642</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeCreateDataSource = 642,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>643</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMailMergeCreateHeaderSource = 643,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditPasteSpecial = 111,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditFind = 112,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditReplace = 117,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>811</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditGoToOld = 811,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>896</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditGoTo = 896,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>872</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogCreateAutoText = 872,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>985</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditAutoText = 985,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditLinks = 124,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditObject = 125,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>392</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogConvertObject = 392,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableToText = 128,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTextToTable = 127,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableInsertTable = 129,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableInsertCells = 130,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableInsertRow = 131,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableDeleteCells = 133,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableSplitCells = 137,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>348</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableFormula = 348,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>563</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableAutoFormat = 563,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>612</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableFormatCell = 612,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>577</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogViewZoom = 577,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>586</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogNewToolbar = 586,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertBreak = 159,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>370</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertFootnote = 370,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertSymbol = 162,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertPicture = 163,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertFile = 164,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertDateTime = 165,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>812</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertNumber = 812,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertField = 166,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>341</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertDatabase = 341,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertMergeField = 167,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertBookmark = 168,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>925</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertHyperlink = 925,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMarkIndexEntry = 169,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>463</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMarkCitation = 463,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>625</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditTOACategory = 625,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>473</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertIndexAndTables = 473,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertIndex = 170,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertTableOfContents = 171,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>442</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogMarkTableOfContentsEntry = 442,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>472</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertTableOfFigures = 472,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>471</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertTableOfAuthorities = 471,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertObject = 172,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>610</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatCallout = 610,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>633</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogDrawSnapToGrid = 633,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>634</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogDrawAlign = 634,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>607</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsEnvelopesAndLabels = 607,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCreateEnvelope = 173,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>489</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCreateLabels = 489,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>503</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsProtectDocument = 503,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>578</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsProtectSection = 578,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>521</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsUnprotectDocument = 521,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>174</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatFont = 174,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatParagraph = 175,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatSectionLayout = 176,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatColumns = 177,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileDocumentLayout = 178,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>685</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileMacPageSetup = 685,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>445</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFilePrintOneCopy = 445,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>444</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileMacPageSetupGX = 444,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>737</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileMacCustomPageSetupGX = 737,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFilePageSetup = 178,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatTabs = 179,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatStyle = 180,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>505</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatStyleGallery = 505,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStyleFont = 181,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStylePara = 182,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>183</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStyleTabs = 183,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>184</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStyleFrame = 184,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>185</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStyleBorders = 185,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>186</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDefineStyleLang = 186,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>187</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatPicture = 187,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>188</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsLanguage = 188,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>189</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatBordersAndShading = 189,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>960</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDrawingObject = 960,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>190</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatFrame = 190,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>488</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatDropCap = 488,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>824</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatBulletsAndNumbering = 824,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>195</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsHyphenation = 195,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>196</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsBulletsNumbers = 196,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>197</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsHighlightChanges = 197,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>506</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAcceptRejectChanges = 506,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>435</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsMergeDocuments = 435,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>198</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCompareDocuments = 198,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>199</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTableSort = 199,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>615</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCustomizeMenuBar = 615,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCustomize = 152,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>432</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCustomizeKeyboard = 432,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>433</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCustomizeMenus = 433,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>723</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogListCommands = 723,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>974</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptions = 974,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>203</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsGeneral = 203,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>206</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAdvancedSettings = 206,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>525</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsCompatibility = 525,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>208</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsPrint = 208,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>209</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsSave = 209,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>211</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsSpellingAndGrammar = 211,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>828</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsSpellingAndGrammar = 828,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>194</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsThesaurus = 194,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>213</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsUserInfo = 213,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>959</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsAutoFormat = 959,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>386</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsTrackChanges = 386,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>224</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsEdit = 224,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>215</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsMacro = 215,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>294</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertPageNumbers = 294,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>298</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatPageNumber = 298,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>373</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogNoteOptions = 373,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>300</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogCopyFile = 300,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatAddrFonts = 103,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>221</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatRetAddrFonts = 221,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>225</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsFileLocations = 225,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>833</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsCreateDirectory = 833,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>331</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogUpdateTOC = 331,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>483</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertFormField = 483,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>353</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormFieldOptions = 353,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>357</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertCaption = 357,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>359</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertAutoCaption = 359,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>402</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertAddCaption = 402,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>358</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertCaptionNumbering = 358,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>367</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertCrossReference = 367,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>631</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsManageFields = 631,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>915</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAutoManager = 915,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>378</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAutoCorrect = 378,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>762</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAutoCorrectExceptions = 762,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>420</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogConnect = 420,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1029</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsBidi = 1029,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>204</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsView = 204,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>583</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogInsertSubdocument = 583,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>624</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileRoutingSlip = 624,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>581</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFontSubstitution = 581,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>732</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditCreatePublisher = 732,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>733</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditSubscribeTo = 733,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>735</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditPublishOptions = 735,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>736</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEditSubscribeOptions = 736,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>739</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsTypography = 739,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>778</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsOptionsAutoFormatAsYouType = 778,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>235</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogControlRun = 235,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>945</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileVersions = 945,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>874</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsAutoSummarize = 874,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1007</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFileSaveVersion = 1007,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>220</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogWindowActivate = 220,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>214</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsMacroRecord = 214,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>197</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogToolsRevisions = 197,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>863</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogEmailOptions = 863,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>898</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogWebOptions = 898,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>983</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFitText = 983,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>986</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogPhoneticGuide = 986,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1160</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogHorizontalInVertical = 1160,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1161</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTwoLinesInOne = 1161,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1162</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatEncloseCharacters = 1162,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>855</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogFormatTheme = 855,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1156</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDialogTCSCTranslator = 1156,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogEditStyle = 120,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableRowHeight = 142,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>143</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableColumnWidth = 143,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>361</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogFormFieldHelp = 361,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>458</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogEditFrame = 458,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>470</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableOfContentsOptions = 470,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>551</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableOfCaptionsOptions = 551,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>570</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogReviewAfmtRevisions = 570,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>784</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsHangulHanjaConversion = 784,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>854</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableWrapping = 854,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>861</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableProperties = 861,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>885</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsGrammarSettings = 885,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>989</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsDictionary = 989,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1074</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogFrameSetProperties = 1074,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1080</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableTableOptions = 1080,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1081</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogTableCellOptions = 1081,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1094</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogIMESetDefault = 1094,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1121</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogConsistencyChecker = 1121,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1395</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsOptionsSmartTag = 1395,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1248</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogFormatStylesCustom = 1248,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1261</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogCSSLinks = 1261,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1324</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogInsertWebComponent = 1324,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1356</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsOptionsEditCopyPaste = 1356,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1361</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogToolsOptionsSecurity = 1361,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1363</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogSearch = 1363,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1381</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogShowRepairs = 1381,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1304</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeFieldMapping = 1304,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1305</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeInsertAddressBlock = 1305,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1306</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeInsertGreetingLine = 1306,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1307</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeInsertFields = 1307,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1308</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeRecipients = 1308,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1326</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeFindRecipient = 1326,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1339</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDialogMailMergeSetDocumentType = 1339,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1460</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogXMLElementAttributes = 1460,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1417</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogSchemaLibrary = 1417,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1469</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogPermission = 1469,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1437</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogMyPermission = 1437,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1425</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogXMLOptions = 1425,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>1427</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdDialogFormattingRestrictions = 1427,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1367</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogLabelOptions = 1367,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1920</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogSourceManager = 1920,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1922</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogCreateSource = 1922,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1482</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogDocumentInspector = 1482,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>1948</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogStyleManagement = 1948,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2120</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogInsertSource = 2120,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2165</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogOMathRecognizedFunctions = 2165,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2348</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogInsertPlaceholder = 2348,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2067</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogBuildingBlockOrganizer = 2067,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2394</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogContentControlProperties = 2394,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2439</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogCompatibilityChecker = 2439,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>2349</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDialogExportAsFixedFormat = 2349,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>1116</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdDialogFileNew2007 = 1116
	}
}