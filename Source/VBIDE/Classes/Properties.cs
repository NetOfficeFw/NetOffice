﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi
{
    /// <summary>
	/// CoClass Properties
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsCoClass)]
    public interface Properties : _Properties
    {
    }
}
