using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
    /// <summary>
    /// DispatchInterface IMsoContactCard 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861128.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public interface IMsoContactCard : _IMsoDispObj
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865180.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        string Address { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864963.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoContactCardAddressType AddressType { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860217.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        NetOffice.OfficeApi.Enums.MsoContactCardType CardType { get; }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860611.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        object Parent { get; }

        #endregion
    }
}
