using System;
using NetOffice;
namespace NetOffice.MSProjectApi.Enums
{
	 /// <summary>
	 /// SupportByVersion MSProject 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("MSProject", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum PjThemeElement
	{
		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>32</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjThemeElementWPBkgd = 32,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>33</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjThemeElementWPText = 33,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>35</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjThemeElementWPTitleBkgdInactive = 35,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>40</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjThemeElementWPCtlBdr = 40,

		 /// <summary>
		 /// SupportByVersion MSProject 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("MSProject", 12,14,15)]
		 pjThemeElementWPGroupline = 47
	}
}