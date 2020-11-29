﻿using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoPatternType"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoPatternType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern5Percent = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern10Percent = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern20Percent = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern25Percent = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern30Percent = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern40Percent = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern50Percent = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern60Percent = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern70Percent = 9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern75Percent = 10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern80Percent = 11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPattern90Percent = 12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDarkHorizontal = 13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDarkVertical = 14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDarkDownwardDiagonal = 15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDarkUpwardDiagonal = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternSmallCheckerBoard = 17,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternTrellis = 18,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLightHorizontal = 19,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLightVertical = 20,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLightDownwardDiagonal = 21,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLightUpwardDiagonal = 22,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternSmallGrid = 23,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDottedDiamond = 24,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternWideDownwardDiagonal = 25,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternWideUpwardDiagonal = 26,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDashedUpwardDiagonal = 27,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDashedDownwardDiagonal = 28,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternNarrowVertical = 29,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternNarrowHorizontal = 30,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDashedVertical = 31,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDashedHorizontal = 32,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLargeConfetti = 33,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLargeGrid = 34,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternHorizontalBrick = 35,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternLargeCheckerBoard = 36,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternSmallConfetti = 37,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternZigZag = 38,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternSolidDiamond = 39,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDiagonalBrick = 40,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternOutlinedDiamond = 41,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternPlaid = 42,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>43</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternSphere = 43,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>44</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternWeave = 44,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>45</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDottedGrid = 45,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>46</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternDivot = 46,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternShingle = 47,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>48</remarks>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPatternWave = 48,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>49</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternHorizontal = 49,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>50</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternVertical = 50,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>51</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternCross = 51,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>52</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternDownwardDiagonal = 52,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>53</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternUpwardDiagonal = 53,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>54</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoPatternDiagonalCross = 54
	}
}