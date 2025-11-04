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
	/// Indicates the assignment method in a <see cref="LabelInfo"/> object.
	/// </summary>
	/// <remarks>
	/// <para>MSDN Online Documentation: <see href="https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/msoassignmentmethod-enumeration-office"/></para>
	/// </remarks>
	[SupportByVersion("Office", 16)]
	[EntityType(EntityType.IsEnum)]
	public enum MsoAssignmentMethod
	{
		/// <summary>
		/// The label is applied by default.
		/// </summary>
		/// <remarks>0</remarks>
		[SupportByVersion("Office", 16)]
		STANDARD = 0,

		/// <summary>
		/// The label was manually selected.
		/// </summary>
		/// <remarks>1</remarks>
		[SupportByVersion("Office", 16)]
		PRIVILEGED = 1,

		/// <summary>
		/// The label is applied automatically.
		/// </summary>
		/// <remarks>2</remarks>
		[SupportByVersion("Office", 16)]
		AUTO = 2
	}
}
