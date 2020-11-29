using System;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Enums
{
	/// <summary>
	/// Specifies for charts, diagrams, or text, the level to which the animation effect will be applied.
	/// The default value is <see cref="msoAnimateLevelNone"/>.
	/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
	/// </summary>
	///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.MsoAnimateByLevel"/> </remarks>
	[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAnimateByLevel
	{
		/// <summary>
		/// Animate level mixed
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>-1</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateLevelMixed = -1,

		/// <summary>
		/// Animate level none
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>0</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateLevelNone = 0,

		/// <summary>
		/// Animate text by all levels
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>1</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextByAllLevels = 1,

		/// <summary>
		/// Animate text by first level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>2</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextByFirstLevel = 2,

		/// <summary>
		/// Animate text by second level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>3</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextBySecondLevel = 3,

		/// <summary>
		/// Animate text by third level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>4</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextByThirdLevel = 4,

		/// <summary>
		/// Animate text by fourth level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>5</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextByFourthLevel = 5,

		/// <summary>
		/// Animate text by fifth level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>6</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateTextByFifthLevel = 6,

		/// <summary>
		/// Animate chart all at once
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>7</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateChartAllAtOnce = 7,

		/// <summary>
		/// Animate chart by category
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>8</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateChartByCategory = 8,

		/// <summary>
		/// Animate chart by category elements
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>9</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateChartByCategoryElements = 9,

		/// <summary>
		/// Animate chart by series
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>10</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateChartBySeries = 10,

		/// <summary>
		/// Animate chart by series elements
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>11</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateChartBySeriesElements = 11,

		/// <summary>
		/// Animate diagram all at once
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>12</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramAllAtOnce = 12,

		/// <summary>
		/// Animate diagram depth by node
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>13</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramDepthByNode = 13,

		/// <summary>
		/// Animate diagram depth by branch
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>14</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramDepthByBranch = 14,

		/// <summary>
		/// Animate diagram breadth by node
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>15</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramBreadthByNode = 15,

		/// <summary>
		/// Animate diagram breadth by level
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>16</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramBreadthByLevel = 16,

		/// <summary>
		/// Animate diagram clockwise
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>17</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramClockwise = 17,

		/// <summary>
		/// Animate diagram clockwise in
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>18</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramClockwiseIn = 18,

		/// <summary>
		/// Animate diagram clockwise out
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>19</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramClockwiseOut = 19,

		/// <summary>
		/// Animate diagram counter-clockwise
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>20</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramCounterClockwise = 20,

		/// <summary>
		/// Animate diagram counter-clockwise in
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>21</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramCounterClockwiseIn = 21,

		/// <summary>
		/// Animate diagram counter-clockwise out
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>22</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramCounterClockwiseOut = 22,

		/// <summary>
		/// Animate diagram in by ring
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>23</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramInByRing = 23,

		/// <summary>
		/// Animate diagram out by ring
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>24</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramOutByRing = 24,

		/// <summary>
		/// Animate diagram up
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>25</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramUp = 25,

		/// <summary>
		/// Animate diagram down
		/// SupportByVersion PowerPoint 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks>26</remarks>
		[SupportByVersion("PowerPoint", 10,11,12,14,15,16)]
		msoAnimateDiagramDown = 26
	}
}