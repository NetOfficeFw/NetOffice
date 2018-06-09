using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    #region Delegates

    #pragma warning disable
    #pragma warning restore

    #endregion


    /// <summary>
    /// CoClass CustomXMLSchemaCollection
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860324.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
	[TypeId("000CDB0D-0000-0000-C000-000000000046")]
    public interface CustomXMLSchemaCollection : _CustomXMLSchemaCollection
    {

    }
}
