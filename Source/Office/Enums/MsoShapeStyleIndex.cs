using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// Represents preset Graphic styles.
	 /// SupportByVersion Office 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoShapeStyleIndex"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoShapeStyleIndex
	{
		 /// <summary>
		 /// A mix of shape styles.
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStyleMixed = -2,

		 /// <summary>
		 /// No shape style.
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStyleNotAPreset = 0,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset1 = 1,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset2 = 2,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset3 = 3,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset4 = 4,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset5 = 5,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset6 = 6,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset7 = 7,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset8 = 8,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset9 = 9,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset10 = 10,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset11 = 11,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset12 = 12,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset13 = 13,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset14 = 14,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset15 = 15,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset16 = 16,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset17 = 17,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset18 = 18,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset19 = 19,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset20 = 20,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset21 = 21,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset22 = 22,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset23 = 23,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset24 = 24,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset25 = 25,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset26 = 26,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset27 = 27,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset28 = 28,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset29 = 29,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset30 = 30,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset31 = 31,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset32 = 32,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset33 = 33,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset34 = 34,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset35 = 35,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset36 = 36,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset37 = 37,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset38 = 38,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset39 = 39,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset40 = 40,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset41 = 41,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>42</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoShapeStylePreset42 = 42,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10001</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset1 = 10001,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10002</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset2 = 10002,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10003</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset3 = 10003,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10004</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset4 = 10004,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10005</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset5 = 10005,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10006</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset6 = 10006,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10007</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset7 = 10007,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10008</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset8 = 10008,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10009</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset9 = 10009,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10010</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset10 = 10010,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10011</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset11 = 10011,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10012</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset12 = 10012,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10013</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset13 = 10013,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10014</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset14 = 10014,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10015</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset15 = 10015,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10016</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset16 = 10016,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10017</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset17 = 10017,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10018</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset18 = 10018,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10019</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset19 = 10019,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10020</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset20 = 10020,

		 /// <summary>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10021</remarks>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoLineStylePreset21 = 10021
	}
}