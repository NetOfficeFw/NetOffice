using System;
using NetOffice;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Office 9, 10, 11, 12, 14
	 /// </summary>
	[SupportByVersionAttribute("Office", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum MsoShapeType
	{
		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>-2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoShapeTypeMixed = -2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>1</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoAutoShape = 1,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>2</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoCallout = 2,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>3</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoChart = 3,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>4</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoComment = 4,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>5</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoFreeform = 5,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>6</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoGroup = 6,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>7</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoEmbeddedOLEObject = 7,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>8</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoFormControl = 8,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>9</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoLine = 9,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>10</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoLinkedOLEObject = 10,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>11</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoLinkedPicture = 11,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>12</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoOLEControlObject = 12,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>13</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoPicture = 13,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>14</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoPlaceholder = 14,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>15</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoTextEffect = 15,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>16</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoMedia = 16,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>17</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoTextBox = 17,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>18</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoScriptAnchor = 18,

		 /// <summary>
		 /// SupportByVersion Office 9, 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>19</remarks>
		 [SupportByVersionAttribute("Office", 9,10,11,12,14)]
		 msoTable = 19,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>20</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoCanvas = 20,

		 /// <summary>
		 /// SupportByVersion Office 10, 11, 12, 14
		 /// </summary>
		 /// <remarks>21</remarks>
		 [SupportByVersionAttribute("Office", 10,11,12,14)]
		 msoDiagram = 21,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>22</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoInk = 22,

		 /// <summary>
		 /// SupportByVersion Office 11, 12, 14
		 /// </summary>
		 /// <remarks>23</remarks>
		 [SupportByVersionAttribute("Office", 11,12,14)]
		 msoInkComment = 23,

		 /// <summary>
		 /// SupportByVersion Office 12, 14
		 /// </summary>
		 /// <remarks>24</remarks>
		 [SupportByVersionAttribute("Office", 12,14)]
		 msoSmartArt = 24,

		 /// <summary>
		 /// SupportByVersion Office 14
		 /// </summary>
		 /// <remarks>25</remarks>
		 [SupportByVersionAttribute("Office", 14)]
		 msoSlicer = 25
	}
}