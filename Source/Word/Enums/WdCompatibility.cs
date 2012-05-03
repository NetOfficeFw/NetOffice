using System;
using NetOffice;
namespace NetOffice.WordApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Word 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum WdCompatibility
	{
		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoTabHangIndent = 1,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoSpaceRaiseLower = 2,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPrintColBlack = 3,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdWrapTrailSpaces = 4,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoColumnBalance = 5,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdConvMailMergeEsc = 6,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSuppressSpBfAfterPgBrk = 7,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSuppressTopSpacing = 8,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdOrigWordTableRules = 9,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdTransparentMetafiles = 10,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdShowBreaksInFrames = 11,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSwapBordersFacingPages = 12,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdLeaveBackslashAlone = 13,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdExpandShiftReturn = 14,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDontULTrailSpace = 15,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDontBalanceSingleByteDoubleByteWidth = 16,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSuppressTopSpacingMac5 = 17,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSpacingInWholePoints = 18,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdPrintBodyTextBeforeHeader = 19,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoLeading = 20,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoSpaceForUL = 21,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdMWSmallCaps = 22,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdNoExtraLineSpacing = 23,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdTruncateFontHeight = 24,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSubFontBySize = 25,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdUsePrinterMetrics = 26,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdWW6BorderRules = 27,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdExactOnTop = 28,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdSuppressBottomSpacing = 29,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdWPSpaceWidth = 30,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdWPJustification = 31,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdLineWrapLikeWord6 = 32,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdShapeLayoutLikeWW8 = 33,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdFootnoteLayoutLikeWW8 = 34,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDontUseHTMLParagraphAutoSpacing = 35,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdDontAdjustLineHeightInTable = 36,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdForgetLastTabAlignment = 37,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdAutospaceLikeWW7 = 38,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdAlignTablesRowByRow = 39,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdLayoutRawTableWidth = 40,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdLayoutTableRowsApart = 41,

		 /// <summary>
		 /// SupportByVersion Word 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersionAttribute("Word", 9,10,11,12,14)]
		 wdUseWord97LineBreakingRules = 42,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDontBreakWrappedTables = 43,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDontSnapTextToGridInTableWithObjects = 44,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdSelectFieldWithFirstOrLastCharacter = 45,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdApplyBreakingRules = 46,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDontWrapTextWithPunctuation = 47,

		 /// <summary>
		 /// SupportByVersion Word 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersionAttribute("Word", 10,11,12,14)]
		 wdDontUseAsianBreakRulesInGrid = 48,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdUseWord2002TableStyleRules = 49,

		 /// <summary>
		 /// SupportByVersion Word 11, 12, 14
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersionAttribute("Word", 11,12,14)]
		 wdGrowAutofit = 50,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdUseNormalStyleForList = 51,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDontUseIndentAsNumberingTabStop = 52,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdFELineBreak11 = 53,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdAllowSpaceOfSameStyleInTable = 54,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdWW11IndentRules = 55,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDontAutofitConstrainedTables = 56,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdAutofitLikeWW11 = 57,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdUnderlineTabInNumList = 58,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdHangulWidthLikeWW11 = 59,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdSplitPgBreakAndParaMark = 60,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDontVertAlignCellWithShape = 61,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDontBreakConstrainedForcedTables = 62,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdDontVertAlignInTextbox = 63,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdWord11KerningPairs = 64,

		 /// <summary>
		 /// SupportByVersion Word 12, 14
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersionAttribute("Word", 12,14)]
		 wdCachedColBalance = 65,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdDisableOTKerning = 66,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdFlipMirrorIndents = 67,

		 /// <summary>
		 /// SupportByVersion Word 14
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Word", 14)]
		 wdDontOverrideTableStyleFontSzAndJustification = 68
	}
}