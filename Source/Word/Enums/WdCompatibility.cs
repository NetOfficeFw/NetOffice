using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838995.aspx </remarks>
	[SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCompatibility
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoTabHangIndent = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoSpaceRaiseLower = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdPrintColBlack = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdWrapTrailSpaces = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoColumnBalance = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdConvMailMergeEsc = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSuppressSpBfAfterPgBrk = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSuppressTopSpacing = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdOrigWordTableRules = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTransparentMetafiles = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdShowBreaksInFrames = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSwapBordersFacingPages = 12,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLeaveBackslashAlone = 13,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdExpandShiftReturn = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDontULTrailSpace = 15,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDontBalanceSingleByteDoubleByteWidth = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSuppressTopSpacingMac5 = 17,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSpacingInWholePoints = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdPrintBodyTextBeforeHeader = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoLeading = 20,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoSpaceForUL = 21,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdMWSmallCaps = 22,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdNoExtraLineSpacing = 23,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdTruncateFontHeight = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSubFontBySize = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdUsePrinterMetrics = 26,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdWW6BorderRules = 27,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdExactOnTop = 28,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdSuppressBottomSpacing = 29,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdWPSpaceWidth = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdWPJustification = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLineWrapLikeWord6 = 32,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdShapeLayoutLikeWW8 = 33,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdFootnoteLayoutLikeWW8 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDontUseHTMLParagraphAutoSpacing = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdDontAdjustLineHeightInTable = 36,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdForgetLastTabAlignment = 37,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdAutospaceLikeWW7 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdAlignTablesRowByRow = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLayoutRawTableWidth = 40,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdLayoutTableRowsApart = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14,15)]
		 wdUseWord97LineBreakingRules = 42,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDontBreakWrappedTables = 43,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDontSnapTextToGridInTableWithObjects = 44,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdSelectFieldWithFirstOrLastCharacter = 45,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdApplyBreakingRules = 46,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDontWrapTextWithPunctuation = 47,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14,15)]
		 wdDontUseAsianBreakRulesInGrid = 48,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdUseWord2002TableStyleRules = 49,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14,15)]
		 wdGrowAutofit = 50,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdUseNormalStyleForList = 51,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDontUseIndentAsNumberingTabStop = 52,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdFELineBreak11 = 53,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdAllowSpaceOfSameStyleInTable = 54,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdWW11IndentRules = 55,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDontAutofitConstrainedTables = 56,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdAutofitLikeWW11 = 57,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdUnderlineTabInNumList = 58,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdHangulWidthLikeWW11 = 59,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdSplitPgBreakAndParaMark = 60,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDontVertAlignCellWithShape = 61,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDontBreakConstrainedForcedTables = 62,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdDontVertAlignInTextbox = 63,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdWord11KerningPairs = 64,

		 /// <summary>
		 /// SupportByVersion Word 12, 14, 15
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Word", 12,14,15)]
		 wdCachedColBalance = 65,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 wdDisableOTKerning = 66,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 wdFlipMirrorIndents = 67,

		 /// <summary>
		 /// SupportByVersion Word 14, 15
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Word", 14,15)]
		 wdDontOverrideTableStyleFontSzAndJustification = 68,

		 /// <summary>
		 /// SupportByVersion Word 15
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersionAttribute("Word", 15)]
		 wdUseWord2010TableStyleRules = 69
	}
}