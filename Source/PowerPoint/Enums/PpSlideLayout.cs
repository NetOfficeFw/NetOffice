using System;
using NetOffice;
namespace NetOffice.PowerPointApi.Enums
{
	 /// <summary>
	 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PpSlideLayout
	{
		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutMixed = -2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTitle = 1,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutText = 2,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTwoColumnText = 3,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTable = 4,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextAndChart = 5,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutChartAndText = 6,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutOrgchart = 7,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutChart = 8,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextAndClipart = 9,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutClipartAndText = 10,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTitleOnly = 11,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutBlank = 12,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextAndObject = 13,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutObjectAndText = 14,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutLargeObject = 15,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutObject = 16,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextAndMediaClip = 17,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutMediaClipAndText = 18,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutObjectOverText = 19,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextOverObject = 20,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTextAndTwoObjects = 21,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTwoObjectsAndText = 22,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutTwoObjectsOverText = 23,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutFourObjects = 24,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutVerticalText = 25,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutClipArtAndVerticalText = 26,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutVerticalTitleAndText = 27,

		 /// <summary>
		 /// SupportByVersion PowerPoint 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersionAttribute("PowerPoint", 9,10,11,12,14)]
		 ppLayoutVerticalTitleAndTextOverChart = 28,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppLayoutTwoObjects = 29,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppLayoutObjectAndTwoObjects = 30,

		 /// <summary>
		 /// SupportByVersion PowerPoint 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersionAttribute("PowerPoint", 10,11,12,14)]
		 ppLayoutTwoObjectsAndObject = 31,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppLayoutCustom = 32,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppLayoutSectionHeader = 33,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppLayoutComparison = 34,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppLayoutContentWithCaption = 35,

		 /// <summary>
		 /// SupportByVersion PowerPoint 12, 14
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersionAttribute("PowerPoint", 12,14)]
		 ppLayoutPictureWithCaption = 36
	}
}