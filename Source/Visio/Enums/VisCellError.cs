using System;
using NetOffice;
namespace NetOffice.VisioApi.Enums
{
	 /// <summary>
	 /// SupportByVersion Visio 11, 12, 14, 15
	 /// </summary>
	[SupportByVersionAttribute("Visio", 11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsEnum)]
	public enum VisCellError
	{
		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>0</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorSuccess = 0,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>39</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorDivideByZero = 39,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>47</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorValue = 47,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>55</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorReference = 55,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>61</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorName = 61,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>68</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorNumber = 68,

		 /// <summary>
		 /// SupportByVersion Visio 11, 12, 14, 15
		 /// </summary>
		 /// <remarks>74</remarks>
		 [SupportByVersionAttribute("Visio", 11,12,14,15)]
		 visErrorNotAvailable = 74
	}
}