using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 11, 12, 14
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff863994(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsEnum)]
	public enum PjThemeElement
	{
		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjThemeElementWPBkgd = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjThemeElementWPText = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjThemeElementWPTitleBkgdInactive = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjThemeElementWPCtlBdr = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 11, 12, 14
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersion("MSProject", 11,12,14)]
		 pjThemeElementWPGroupline = 47
	}
}