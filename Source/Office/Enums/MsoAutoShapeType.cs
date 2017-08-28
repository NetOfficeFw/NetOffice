using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862770.aspx </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAutoShapeType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRectangle = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeParallelogram = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeTrapezoid = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDiamond = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRoundedRectangle = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeOctagon = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeIsoscelesTriangle = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRightTriangle = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeOval = 9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeHexagon = 10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCross = 11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRegularPentagon = 12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCan = 13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCube = 14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeBevel = 15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFoldedCorner = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeSmileyFace = 17,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDonut = 18,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeNoSymbol = 19,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeBlockArc = 20,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeHeart = 21,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLightningBolt = 22,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeSun = 23,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeMoon = 24,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeArc = 25,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDoubleBracket = 26,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDoubleBrace = 27,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapePlaque = 28,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftBracket = 29,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRightBracket = 30,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftBrace = 31,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRightBrace = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRightArrow = 33,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftArrow = 34,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUpArrow = 35,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDownArrow = 36,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftRightArrow = 37,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUpDownArrow = 38,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeQuadArrow = 39,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftRightUpArrow = 40,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeBentArrow = 41,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUTurnArrow = 42,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftUpArrow = 43,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeBentUpArrow = 44,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedRightArrow = 45,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedLeftArrow = 46,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedUpArrow = 47,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedDownArrow = 48,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeStripedRightArrow = 49,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeNotchedRightArrow = 50,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapePentagon = 51,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeChevron = 52,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRightArrowCallout = 53,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftArrowCallout = 54,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUpArrowCallout = 55,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>56</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDownArrowCallout = 56,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>57</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLeftRightArrowCallout = 57,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>58</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUpDownArrowCallout = 58,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>59</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeQuadArrowCallout = 59,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>60</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCircularArrow = 60,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartProcess = 61,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>62</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartAlternateProcess = 62,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>63</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartDecision = 63,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>64</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartData = 64,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>65</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartPredefinedProcess = 65,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>66</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartInternalStorage = 66,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>67</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartDocument = 67,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartMultidocument = 68,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>69</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartTerminator = 69,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>70</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartPreparation = 70,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>71</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartManualInput = 71,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>72</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartManualOperation = 72,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>73</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartConnector = 73,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartOffpageConnector = 74,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>75</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartCard = 75,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>76</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartPunchedTape = 76,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>77</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartSummingJunction = 77,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>78</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartOr = 78,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>79</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartCollate = 79,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>80</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartSort = 80,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>81</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartExtract = 81,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>82</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartMerge = 82,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>83</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartStoredData = 83,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>84</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartDelay = 84,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>85</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartSequentialAccessStorage = 85,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>86</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartMagneticDisk = 86,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>87</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartDirectAccessStorage = 87,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>88</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeFlowchartDisplay = 88,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>89</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeExplosion1 = 89,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>90</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeExplosion2 = 90,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>91</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape4pointStar = 91,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>92</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape5pointStar = 92,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>93</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape8pointStar = 93,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>94</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape16pointStar = 94,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>95</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape24pointStar = 95,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>96</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShape32pointStar = 96,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>97</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeUpRibbon = 97,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>98</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDownRibbon = 98,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>99</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedUpRibbon = 99,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>100</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCurvedDownRibbon = 100,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>101</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeVerticalScroll = 101,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>102</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeHorizontalScroll = 102,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>103</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeWave = 103,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>104</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeDoubleWave = 104,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>105</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRectangularCallout = 105,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>106</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeRoundedRectangularCallout = 106,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>107</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeOvalCallout = 107,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>108</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeCloudCallout = 108,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>109</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout1 = 109,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>110</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout2 = 110,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>111</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout3 = 111,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>112</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout4 = 112,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>113</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout1AccentBar = 113,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>114</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout2AccentBar = 114,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>115</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout3AccentBar = 115,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>116</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout4AccentBar = 116,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>117</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout1NoBorder = 117,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>118</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout2NoBorder = 118,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>119</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout3NoBorder = 119,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>120</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout4NoBorder = 120,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>121</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout1BorderandAccentBar = 121,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>122</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout2BorderandAccentBar = 122,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>123</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout3BorderandAccentBar = 123,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>124</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeLineCallout4BorderandAccentBar = 124,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>125</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonCustom = 125,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>126</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonHome = 126,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>127</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonHelp = 127,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>128</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonInformation = 128,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>129</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonBackorPrevious = 129,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>130</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonForwardorNext = 130,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>131</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonBeginning = 131,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>132</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonEnd = 132,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>133</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonReturn = 133,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>134</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonDocument = 134,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>135</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonSound = 135,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>136</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeActionButtonMovie = 136,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>137</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeBalloon = 137,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>138</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeNotPrimitive = 138,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>139</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeFlowchartOfflineStorage = 139,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>140</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeLeftRightRibbon = 140,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>141</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeDiagonalStripe = 141,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>142</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapePie = 142,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>143</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeNonIsoscelesTrapezoid = 143,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>144</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeDecagon = 144,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>145</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeHeptagon = 145,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>146</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeDodecagon = 146,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>147</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShape6pointStar = 147,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>148</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShape7pointStar = 148,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>149</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShape10pointStar = 149,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>150</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShape12pointStar = 150,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>151</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeRound1Rectangle = 151,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>152</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeRound2SameRectangle = 152,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>153</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeRound2DiagRectangle = 153,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>154</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSnipRoundRectangle = 154,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>155</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSnip1Rectangle = 155,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>156</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSnip2SameRectangle = 156,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>157</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSnip2DiagRectangle = 157,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>158</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeFrame = 158,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>159</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeHalfFrame = 159,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>160</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeTear = 160,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>161</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeChord = 161,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>162</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeCorner = 162,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>163</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathPlus = 163,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>164</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathMinus = 164,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>165</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathMultiply = 165,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>166</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathDivide = 166,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>167</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathEqual = 167,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>168</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeMathNotEqual = 168,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>169</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeCornerTabs = 169,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>170</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSquareTabs = 170,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>171</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapePlaqueTabs = 171,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>172</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeGear6 = 172,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>173</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeGear9 = 173,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>174</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeFunnel = 174,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>175</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapePieWedge = 175,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>176</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeLeftCircularArrow = 176,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>177</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeLeftRightCircularArrow = 177,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>178</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeSwooshArrow = 178,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>179</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeCloud = 179,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>180</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeChartX = 180,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>181</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeChartStar = 181,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>182</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeChartPlus = 182,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>183</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeLineInverse = 183
	}
}