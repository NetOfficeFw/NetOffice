using System;
using NetOffice;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Outlook", 12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum OlSearchScope
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14,15)]
		 olSearchScopeCurrentFolder = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Outlook", 12,14,15)]
		 olSearchScopeAllFolders = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olSearchScopeAllOutlookItems = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Outlook", 14,15)]
		 olSearchScopeSubfolders = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Outlook", 15)]
		 olSearchScopeCurrentStore = 4
	}
}