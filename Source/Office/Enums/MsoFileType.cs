using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoFileType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeAllFiles = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeOfficeFiles = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeWordDocuments = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeExcelWorkbooks = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypePowerPointPresentations = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeBinders = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeDatabases = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14,15)]
		 msoFileTypeTemplates = 8,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeOutlookItems = 9,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeMailItem = 10,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeCalendarItem = 11,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeContactItem = 12,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeNoteItem = 13,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeJournalItem = 14,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeTaskItem = 15,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypePhotoDrawFiles = 16,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeDataConnectionFiles = 17,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypePublisherFiles = 18,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeProjectFiles = 19,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeDocumentImagingFiles = 20,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeVisioFiles = 21,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeDesignerFiles = 22,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14,15)]
		 msoFileTypeWebPages = 23
	}
}