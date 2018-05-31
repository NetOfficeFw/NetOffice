using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
    /// <summary>
    /// Interface IHyperlink 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public interface IHyperlink : ICOMObject
    {
        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Application Application { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        object Parent { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Name { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Range Range { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        NetOffice.ExcelApi.Shape Shape { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string SubAddress { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string Address { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Type { get; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string EmailSubject { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string ScreenTip { get; set; }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        string TextToDisplay { get; set; }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 AddToFavorites();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Delete();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        /// <param name="headerInfo">optional object headerInfo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow(object newWindow, object addHistory, object extraInfo, object method, object headerInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow();

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow(object newWindow);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow(object newWindow, object addHistory);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow(object newWindow, object addHistory, object extraInfo);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="newWindow">optional object newWindow</param>
        /// <param name="addHistory">optional object addHistory</param>
        /// <param name="extraInfo">optional object extraInfo</param>
        /// <param name="method">optional object method</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 Follow(object newWindow, object addHistory, object extraInfo, object method);

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="filename">string filename</param>
        /// <param name="editNow">bool editNow</param>
        /// <param name="overwrite">bool overwrite</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        Int32 CreateNewDocument(string filename, bool editNow, bool overwrite);

        #endregion
    }
}
