using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.AccessApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff845524.aspx </remarks>
	[SupportByVersion("Access", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum AcPrintPaperSize
	{
		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSLetter = 1,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSLetterSmall = 2,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSTabloid = 3,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSLedger = 4,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSLegal = 5,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSStatement = 6,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSExecutive = 7,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSA3 = 8,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSA4 = 9,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSA4Small = 10,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSA5 = 11,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSB4 = 12,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSB5 = 13,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSFolio = 14,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSQuarto = 15,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPS10x14 = 16,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPS11x17 = 17,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSNote = 18,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnv9 = 19,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnv10 = 20,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnv11 = 21,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnv12 = 22,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnv14 = 23,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSCSheet = 24,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSDSheet = 25,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSESheet = 26,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvDL = 27,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvC3 = 29,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvC4 = 30,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvC5 = 28,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>31</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvC6 = 31,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvC65 = 32,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvB4 = 33,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>34</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvB5 = 34,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvB6 = 35,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>36</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvItaly = 36,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>37</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvMonarch = 37,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>38</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSEnvPersonal = 38,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSFanfoldUS = 39,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSFanfoldStdGerman = 40,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>41</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSFanfoldLglGerman = 41,

		 /// <summary>
		 /// SupportByVersion Access 10, 11, 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>256</remarks>
		 [SupportByVersion("Access", 10,11,12,14,15,16)]
		 acPRPSUser = 256
	}
}