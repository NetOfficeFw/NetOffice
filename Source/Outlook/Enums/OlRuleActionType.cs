using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OutlookApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Outlook 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff869344.aspx </remarks>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum OlRuleActionType
	{
		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionUnknown = 0,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionMoveToFolder = 1,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionAssignToCategory = 2,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionDelete = 3,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionDeletePermanently = 4,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionCopyToFolder = 5,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionForward = 6,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionForwardAsAttachment = 7,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionRedirect = 8,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionServerReply = 9,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionTemplate = 10,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionFlagForActionInDays = 11,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionFlagColor = 12,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionFlagClear = 13,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionImportance = 14,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionSensitivity = 15,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionPrint = 16,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionPlaySound = 17,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionStartApplication = 18,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionMarkRead = 19,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionRunScript = 20,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionStop = 21,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionCustomAction = 22,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionNewItemAlert = 23,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionDesktopAlert = 24,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionNotifyRead = 25,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>26</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionNotifyDelivery = 26,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>27</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionCcMessage = 27,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>28</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionDefer = 28,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>29</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionMarkAsTask = 29,

		 /// <summary>
		 /// SupportByVersion Outlook 12, 14, 15, 16
		 /// </summary>
		 /// <remarks>30</remarks>
		 [SupportByVersion("Outlook", 12,14,15,16)]
		 olRuleActionClearCategories = 30
	}
}