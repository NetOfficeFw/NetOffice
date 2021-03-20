using System;
using NetOffice;
using NetOffice.Attributes;
namespace NetOffice.OfficeApi.Enums
{
	 /// <summary>
	 /// Specifies the type of a shape or range of shapes.
	 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
	 /// </summary>
	 ///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.MsoShapeType"/> </remarks>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoShapeType
	{
		 /// <summary>
		 /// Mixed shape type
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>-2</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoShapeTypeMixed = -2,

		 /// <summary>
		 /// AutoShape
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>1</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoAutoShape = 1,

		 /// <summary>
		 /// Callout
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>2</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoCallout = 2,

		 /// <summary>
		 /// Chart
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>3</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoChart = 3,

		 /// <summary>
		 /// Comment
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>4</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoComment = 4,

		 /// <summary>
		 /// Freeform
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>5</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoFreeform = 5,

		 /// <summary>
		 /// Group
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>6</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoGroup = 6,

		 /// <summary>
		 /// Embedded OLE object
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>7</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoEmbeddedOLEObject = 7,

		 /// <summary>
		 /// Form control
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>8</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoFormControl = 8,

		 /// <summary>
		 /// Line
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>9</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLine = 9,

		 /// <summary>
		 /// Linked OLE object
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>10</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLinkedOLEObject = 10,

		 /// <summary>
		 /// Linked picture
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>11</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoLinkedPicture = 11,

		 /// <summary>
		 /// OLE control object
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>12</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoOLEControlObject = 12,

		 /// <summary>
		 /// Picture
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>13</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPicture = 13,

		 /// <summary>
		 /// Placeholder
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>14</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoPlaceholder = 14,

		 /// <summary>
		 /// Text effect
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>15</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoTextEffect = 15,

		 /// <summary>
		 /// Media
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>16</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoMedia = 16,

		 /// <summary>
		 /// Text box
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>17</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoTextBox = 17,

		 /// <summary>
		 /// Script anchor
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>18</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoScriptAnchor = 18,

		 /// <summary>
		 /// Table
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>19</value>
		 [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		 msoTable = 19,

		 /// <summary>
		 /// Canvas
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>20</value>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoCanvas = 20,

		 /// <summary>
		 /// Diagram
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 10, 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>21</value>
		 [SupportByVersion("Office", 10,11,12,14,15,16)]
		 msoDiagram = 21,

		 /// <summary>
		 /// Ink
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>22</value>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoInk = 22,

		 /// <summary>
		 /// Ink comment
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 11, 12, 14, 15, 16
		 /// </remarks>
		 /// <value>23</value>
		 [SupportByVersion("Office", 11,12,14,15,16)]
		 msoInkComment = 23,

		 /// <summary>
		 /// SmartArt graphic
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 12, 14, 15, 16
		 /// </remarks>
		 /// <value>24</value>
		 [SupportByVersion("Office", 12,14,15,16)]
		 msoSmartArt = 24,

		 /// <summary>
		 /// Slicer
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 14, 15, 16
		 /// </remarks>
		 /// <value>25</value>
		 [SupportByVersion("Office", 14,15,16)]
		 msoSlicer = 25,

		 /// <summary>
		 /// Web video
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 msoWebVideo = 26,

		 /// <summary>
		 /// Content Office Add-in
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 msoContentApp = 27,

		 /// <summary>
		 /// Graphic
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 msoGraphic = 28,

		 /// <summary>
		 /// Linked graphic
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 msoLinkedGraphic = 29,

		 /// <summary>
		 /// 3D model
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 mso3DModel = 30,

		 /// <summary>
		 /// Linked 3D model
		 /// </summary>
		 /// <remarks>
		 /// SupportByVersion Office 15,16
		 /// </remarks>
		 /// <value>26</value>
		 [SupportByVersion("Office", 15, 16)]
		 msoLinked3DModel = 31,
	}
}