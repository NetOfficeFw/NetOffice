// Copyright 2025 Cisco Systems, Inc. All rights reserved.
// Licensed under MIT-style license (see LICENSE.txt file).
//
// Generated code file by Claude Haiku 4.5
//

using System;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi.Enums
{
	/// <summary>
	/// SupportByVersion Office 16
	/// </summary>
	/// <remarks>
	/// Specifies the assignment method for a sensitivity label.
	/// <para>MSDN Online Documentation: <see href="https://learn.microsoft.com/en-us/office/vba/api/office.msoassignmentmethod"/></para>
	/// </remarks>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAssignmentMethod
	{
		/// <summary>
		/// SupportByVersion Office 16
		/// </summary>
		/// <remarks>0</remarks>
		[SupportByVersion("Office", 16)]
		STANDARD = 0,

		/// <summary>
		/// SupportByVersion Office 16
		/// </summary>
		/// <remarks>1</remarks>
		[SupportByVersion("Office", 16)]
		PRIVILEGED = 1,

		/// <summary>
		/// SupportByVersion Office 16
		/// </summary>
		/// <remarks>2</remarks>
		[SupportByVersion("Office", 16)]
		AUTO = 2
	}
}
