using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
    /// <summary>
    /// DispatchInterface IHTMLInputTextElement 
    /// SupportByVersion MSHTML, 4
    /// </summary>
    [SupportByVersion("MSHTML", 4)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
	[TypeId("3050F2A6-98B5-11CF-BB82-00AA00BDCE0B")]
    public interface IHTMLInputTextElement : IHTMLInputElement2
    {
        /// <summary>
        /// SupportByVersion MSHTML 4
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSHTML", 4)]
        new object status { get; set; }
    }
}
