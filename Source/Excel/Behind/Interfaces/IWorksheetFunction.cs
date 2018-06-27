using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IWorksheetFunction 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public partial class IWorksheetFunction : COMObject, NetOffice.ExcelApi.IWorksheetFunction
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.ExcelApi.IWorksheetFunction);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        private static Type _type;

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IWorksheetFunction);
                return _type;
            }
        }

        #endregion

        #region Ctor
        
        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public IWorksheetFunction() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsNA(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsNA", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsError(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsError", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Dollar(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dollar", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Dollar(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dollar", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Fixed(Double arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Fixed", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Fixed(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Fixed", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Fixed(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Fixed", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pi()
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pi");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ln(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ln", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Log10(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Log10", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Round(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Round", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Lookup(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Lookup", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Lookup(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Lookup", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Index(object arg1, Double arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Index", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Index(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Index", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Index(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Index", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Rept(string arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Rept", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DCount(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DCount", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DSum(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DSum", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DAverage(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DAverage", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DMin(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DMin", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DMax(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DMax", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DStDev(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DStDev", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DVar(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DVar", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">string arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Text(object arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Text", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LinEst(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinEst", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LinEst(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinEst", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LinEst(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinEst", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LinEst(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LinEst", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Trend(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trend", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Trend(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trend", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Trend(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trend", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Trend(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trend", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LogEst(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LogEst", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LogEst(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LogEst", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LogEst(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LogEst", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object LogEst(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LogEst", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Growth(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Growth", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Growth(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Growth", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Growth(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Growth", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Growth(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Growth", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pv(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pv", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pv(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pv", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Fv(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Fv", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Fv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Fv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Fv(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Fv", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NPer(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NPer", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NPer(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NPer", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NPer(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NPer", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pmt(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pmt", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pmt(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pmt", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pmt(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pmt", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rate(Double arg1, Double arg2, Double arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rate", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rate(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rate", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rate(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rate", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rate(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rate", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double MIrr(object arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "MIrr", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Irr(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Irr", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Irr(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Irr", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Match(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Match", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Match(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Match", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Weekday(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Weekday", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Weekday(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Weekday", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Search(string arg1, string arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Search", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Search(string arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Search", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Transpose(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Transpose", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Atan2(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Atan2", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Asin(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Asin", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Acos(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Acos", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object HLookup(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "HLookup", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object HLookup(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "HLookup", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object VLookup(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "VLookup", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object VLookup(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "VLookup", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Log(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Log", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Log(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Log", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Proper(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Proper", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Trim(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Trim", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">string arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Replace(string arg1, Double arg2, Double arg3, string arg4)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Replace", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Substitute(string arg1, string arg2, string arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Substitute", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Substitute(string arg1, string arg2, string arg3)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Substitute", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Find(string arg1, string arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Find", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Find(string arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Find", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsErr(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsErr", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsText(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsText", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsNumber(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsNumber", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Sln(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Sln", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Syd(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Syd", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ddb(Double arg1, Double arg2, Double arg3, Double arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ddb", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ddb(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ddb", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Clean(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Clean", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double MDeterm(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "MDeterm", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object MInverse(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "MInverse", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object MMult(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "MMult", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ipmt", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ipmt", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ipmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ipmt", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ppmt", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ppmt", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ppmt(Double arg1, Double arg2, Double arg3, Double arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ppmt", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Fact(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Fact", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DProduct(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DProduct", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsNonText(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsNonText", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DStDevP(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DStDevP", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DVarP(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DVarP", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public bool IsLogical(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsLogical", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double DCountA(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DCountA", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string USDollar(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "USDollar", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FindB(string arg1, string arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FindB", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FindB(string arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FindB", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SearchB(string arg1, string arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SearchB", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SearchB(string arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SearchB", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">string arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string ReplaceB(string arg1, Double arg2, Double arg3, string arg4)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ReplaceB", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double RoundUp(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RoundUp", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double RoundDown(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RoundDown", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rank(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Rank(Double arg1, NetOffice.ExcelApi.Range arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Days360(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Days360", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Days360(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Days360", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        /// <param name="arg6">optional object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Vdb", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Vdb", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">Double arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Vdb(Double arg1, Double arg2, Double arg3, Double arg4, Double arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Vdb", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Sinh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Sinh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Cosh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Cosh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Tanh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Tanh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Asinh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Asinh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Acosh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Acosh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Atanh(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Atanh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object DGet(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DGet", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Db(Double arg1, Double arg2, Double arg3, Double arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Db", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Db(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Db", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public object Frequency(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Frequency", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaDist(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaDist", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaDist(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaDist(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaDist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double GammaLn(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GammaLn", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaInv(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaInv", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaInv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaInv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BetaInv(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BetaInv", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double BinomDist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BinomDist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ChiDist(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiDist", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ChiInv(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiInv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Combin(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Combin", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Confidence(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Confidence", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double CritBinom(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CritBinom", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Even(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Even", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ExponDist(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ExponDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FDist(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FInv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FInv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Fisher(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Fisher", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FisherInv(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FisherInv", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Floor(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double GammaDist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GammaDist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double GammaInv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GammaInv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ceiling(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double HypGeomDist(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "HypGeomDist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double LogNormDist(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "LogNormDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double LogInv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "LogInv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NegBinomDist(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NegBinomDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NormDist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NormDist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NormSDist(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NormSDist", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NormInv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NormInv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double NormSInv(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NormSInv", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Standardize(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Standardize", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Odd(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Odd", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Permut(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Permut", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Poisson(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Poisson", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double TDist(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TDist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Weibull(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Weibull", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SumXMY2(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SumXMY2", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SumX2MY2(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SumX2MY2", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SumX2PY2(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SumX2PY2", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ChiTest(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiTest", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Correl(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Correl", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Covar(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Covar", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Forecast(Double arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Forecast", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double FTest(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FTest", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Intercept(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Intercept", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Pearson(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Pearson", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double RSq(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RSq", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double StEyx(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "StEyx", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Slope(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Slope", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double TTest(object arg1, object arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TTest", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Prob(object arg1, object arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Prob", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Prob(object arg1, object arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Prob", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ZTest(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ZTest", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double ZTest(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ZTest", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Large(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Large", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Small(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Small", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Quartile(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Quartile", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Percentile(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Percentile", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double PercentRank(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double PercentRank(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double TrimMean(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TrimMean", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double TInv(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TInv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Power(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Power", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Radians(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Radians", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Degrees(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Degrees", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SumIf(NetOffice.ExcelApi.Range arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SumIf", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double SumIf(NetOffice.ExcelApi.Range arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SumIf", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double CountIf(NetOffice.ExcelApi.Range arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CountIf", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double CountBlank(NetOffice.ExcelApi.Range arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CountBlank", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Double Ispmt(Double arg1, Double arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ispmt", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Roman(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Roman", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Roman(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Roman", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Asc(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Asc", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Dbcs(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dbcs", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string Phonetic(NetOffice.ExcelApi.Range arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Phonetic", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string BahtText(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "BahtText", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string ThaiDayOfWeek(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ThaiDayOfWeek", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string ThaiDigit(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ThaiDigit", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string ThaiMonthOfYear(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ThaiMonthOfYear", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string ThaiNumSound(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ThaiNumSound", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public string ThaiNumString(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ThaiNumString", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Double ThaiStringLength(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ThaiStringLength", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public bool IsThaiDigit(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsThaiDigit", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Double RoundBahtDown(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RoundBahtDown", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Double RoundBahtUp(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RoundBahtUp", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public Double ThaiYear(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ThaiYear", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Hex2Bin(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Hex2Bin", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Hex2Bin(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Hex2Bin", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Hex2Dec(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Hex2Dec", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Hex2Oct(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Hex2Oct", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Hex2Oct(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Hex2Oct", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Bin(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Bin", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Bin(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Bin", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Hex(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Hex", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Hex(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Hex", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Oct(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Oct", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Dec2Oct(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Dec2Oct", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Oct2Bin(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Oct2Bin", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Oct2Bin(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Oct2Bin", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Oct2Hex(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Oct2Hex", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Oct2Hex(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Oct2Hex", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Oct2Dec(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Oct2Dec", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Bin2Dec(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Bin2Dec", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Bin2Oct(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Bin2Oct", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Bin2Oct(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Bin2Oct", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Bin2Hex(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Bin2Hex", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Bin2Hex(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Bin2Hex", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImSub(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSub", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImDiv(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImDiv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImPower(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImPower", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImAbs(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImAbs", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImSqrt(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSqrt", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImLn(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImLn", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImLog2(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImLog2", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImLog10(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImLog10", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImSin(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSin", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImCos(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImCos", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImExp(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImExp", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImArgument(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImArgument", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string ImConjugate(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImConjugate", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Imaginary(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Imaginary", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double ImReal(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ImReal", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Complex(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Complex", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public string Complex(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Complex", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double SeriesSum(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SeriesSum", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double FactDouble(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FactDouble", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double SqrtPi(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "SqrtPi", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Quotient(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Quotient", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Delta(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Delta", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Delta(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Delta", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double GeStep(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GeStep", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double GeStep(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GeStep", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool IsEven(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsEven", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public bool IsOdd(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsOdd", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double MRound(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "MRound", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Erf(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Erf", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Erf(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Erf", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double ErfC(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ErfC", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double BesselJ(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BesselJ", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double BesselK(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BesselK", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double BesselY(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BesselY", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double BesselI(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "BesselI", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Xirr(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Xirr", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Xirr(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Xirr", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Xnpv(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Xnpv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double PriceMat(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PriceMat", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double PriceMat(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PriceMat", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YieldMat(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YieldMat", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YieldMat(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YieldMat", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double IntRate(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "IntRate", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double IntRate(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "IntRate", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Received(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Received", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Received(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Received", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Disc(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Disc", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Disc(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Disc", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double PriceDisc(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PriceDisc", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double PriceDisc(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PriceDisc", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YieldDisc(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YieldDisc", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YieldDisc(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YieldDisc", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillEq(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillEq", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillEq(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillEq", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillPrice(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillPrice", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillPrice(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillPrice", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillYield(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillYield", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double TBillYield(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "TBillYield", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Price(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Price", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Price(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Price", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double DollarDe(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DollarDe", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double DollarFr(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "DollarFr", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Nominal(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Nominal", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Effect(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Effect", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CumPrinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CumPrinc", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CumIPmt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CumIPmt", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double EDate(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EDate", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double EoMonth(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EoMonth", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YearFrac(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YearFrac", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double YearFrac(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "YearFrac", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDayBs(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDayBs", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDayBs(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDayBs", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDays(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDays", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDays(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDays", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDaysNc(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDaysNc", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupDaysNc(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupDaysNc", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupNcd(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupNcd", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupNcd(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupNcd", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupNum(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupNum", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupNum(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupNum", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupPcd(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupPcd", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double CoupPcd(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "CoupPcd", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Duration(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Duration", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Duration(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Duration", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double MDuration(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "MDuration", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double MDuration(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "MDuration", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddLPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddLPrice", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddLPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddLPrice", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">optional object arg8</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddLYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddLYield", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddLYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddLYield", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddFPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddFPrice", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddFPrice(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddFPrice", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        /// <param name="arg9">optional object arg9</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddFYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8, object arg9)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddFYield", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">object arg7</param>
        /// <param name="arg8">object arg8</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double OddFYield(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7, object arg8)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "OddFYield", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double RandBetween(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "RandBetween", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double WeekNum(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WeekNum", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double WeekNum(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WeekNum", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AmorDegrc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AmorDegrc", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AmorDegrc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AmorDegrc", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AmorLinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AmorLinc", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AmorLinc(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AmorLinc", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double Convert(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Convert", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        /// <param name="arg7">optional object arg7</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AccrInt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6, object arg7)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AccrInt", new object[] { arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">object arg5</param>
        /// <param name="arg6">object arg6</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AccrInt(object arg1, object arg2, object arg3, object arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AccrInt", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AccrIntM(object arg1, object arg2, object arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AccrIntM", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">object arg3</param>
        /// <param name="arg4">object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double AccrIntM(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "AccrIntM", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double WorkDay(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WorkDay", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double WorkDay(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WorkDay", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double NetworkDays(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NetworkDays", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double NetworkDays(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NetworkDays", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public Double FVSchedule(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "FVSchedule", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public object IfError(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "IfError", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Confidence_Norm(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Confidence_Norm", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Confidence_T(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Confidence_T", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ChiSq_Test(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiSq_Test", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double F_Test(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "F_Test", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Covariance_P(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Covariance_P", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Covariance_S(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Covariance_S", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Expon_Dist(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Expon_Dist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Gamma_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Gamma_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Gamma_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Gamma_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Norm_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Norm_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Norm_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Norm_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Percentile_Exc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Percentile_Exc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Percentile_Inc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Percentile_Inc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double PercentRank_Exc(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank_Exc", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double PercentRank_Exc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank_Exc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double PercentRank_Inc(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank_Inc", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double PercentRank_Inc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PercentRank_Inc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Poisson_Dist(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Poisson_Dist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Quartile_Exc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Quartile_Exc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Quartile_Inc(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Quartile_Inc", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Rank_Avg(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank_Avg", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Rank_Avg(Double arg1, NetOffice.ExcelApi.Range arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank_Avg", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Rank_Eq(Double arg1, NetOffice.ExcelApi.Range arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank_Eq", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">NetOffice.ExcelApi.Range arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Rank_Eq(Double arg1, NetOffice.ExcelApi.Range arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rank_Eq", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Dist(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Dist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Dist_2T(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Dist_2T", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Dist_RT(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Dist_RT", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Inv(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Inv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Inv_2T(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Inv_2T", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Weibull_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Weibull_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double NetworkDays_Intl(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NetworkDays_Intl", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double NetworkDays_Intl(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NetworkDays_Intl", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double NetworkDays_Intl(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NetworkDays_Intl", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double WorkDay_Intl(object arg1, object arg2, object arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WorkDay_Intl", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double WorkDay_Intl(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WorkDay_Intl", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double WorkDay_Intl(object arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "WorkDay_Intl", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ISO_Ceiling(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ISO_Ceiling", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ISO_Ceiling(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ISO_Ceiling", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Dummy21(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Dummy21", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        /// <param name="arg5">optional object arg5</param>
        /// <param name="arg6">optional object arg6</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4, object arg5, object arg6)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Dist", new object[] { arg1, arg2, arg3, arg4, arg5, arg6 });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Dist(Double arg1, Double arg2, Double arg3, bool arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Dist", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        /// <param name="arg5">optional object arg5</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Inv(Double arg1, Double arg2, Double arg3, object arg4, object arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Inv", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Beta_Inv(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Beta_Inv", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">bool arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ChiSq_Dist(Double arg1, Double arg2, bool arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiSq_Dist", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ChiSq_Dist_RT(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiSq_Dist_RT", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ChiSq_Inv(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiSq_Inv", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ChiSq_Inv_RT(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ChiSq_Inv_RT", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double F_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "F_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double F_Dist_RT(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "F_Dist_RT", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double F_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "F_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double F_Inv_RT(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "F_Inv_RT", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        /// <param name="arg5">bool arg5</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double HypGeom_Dist(Double arg1, Double arg2, Double arg3, Double arg4, bool arg5)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "HypGeom_Dist", new object[] { arg1, arg2, arg3, arg4, arg5 });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double LogNorm_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "LogNorm_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double LogNorm_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "LogNorm_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double NegBinom_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NegBinom_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">bool arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Norm_S_Dist(Double arg1, bool arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Norm_S_Dist", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Norm_S_Inv(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Norm_S_Inv", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">Double arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double T_Test(object arg1, object arg2, Double arg3, Double arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "T_Test", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Z_Test(object arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Z_Test", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Z_Test(object arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Z_Test", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">bool arg4</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Binom_Dist(Double arg1, Double arg2, Double arg3, bool arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Binom_Dist", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Binom_Inv(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Binom_Inv", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Erf_Precise(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Erf_Precise", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double ErfC_Precise(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "ErfC_Precise", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double GammaLn_Precise(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "GammaLn_Precise", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Ceiling_Precise(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling_Precise", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Ceiling_Precise(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling_Precise", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Floor_Precise(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor_Precise", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public Double Floor_Precise(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor_Precise", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Acot(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Acot", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Acoth(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Acoth", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Cot(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Cot", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Coth(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Coth", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Csc(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Csc", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Csch(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Csch", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Sec(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Sec", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Sech(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Sech", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImCot(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImCot", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImTan(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImTan", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImCsc(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImCsc", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImCsch(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImCsch", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImSec(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSec", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImSech(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSech", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Bitand(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Bitand", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Bitor(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Bitor", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Bitxor(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Bitxor", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Bitlshift(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Bitlshift", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Bitrshift(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Bitrshift", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Combina(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Combina", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Permutationa(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Permutationa", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double PDuration(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "PDuration", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public string Base(Double arg1, Double arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Base", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public string Base(Double arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Base", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">Double arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Decimal(string arg1, Double arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Decimal", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Days(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Days", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        /// <param name="arg4">optional object arg4</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Binom_Dist_Range(Double arg1, Double arg2, Double arg3, object arg4)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Binom_Dist_Range", arg1, arg2, arg3, arg4);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double Binom_Dist_Range(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Binom_Dist_Range", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Gamma(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Gamma", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Gauss(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Gauss", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Phi(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Phi", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">Double arg2</param>
        /// <param name="arg3">Double arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Rri(Double arg1, Double arg2, Double arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Rri", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string Unichar(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "Unichar", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Unicode(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Unicode", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public object Munit(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Munit", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Arabic(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Arabic", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double IsoWeekNum(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "IsoWeekNum", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double IsoWeekNum(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "IsoWeekNum", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        /// <param name="arg3">string arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double NumberValue(string arg1, string arg2, string arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "NumberValue", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">NetOffice.ExcelApi.Range arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public bool IsFormula(NetOffice.ExcelApi.Range arg1)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsFormula", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        /// <param name="arg2">object arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public object IfNa(object arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "IfNa", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Ceiling_Math(Double arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling_Math", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double Ceiling_Math(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling_Math", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double Ceiling_Math(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Ceiling_Math", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        /// <param name="arg3">optional object arg3</param>
        [SupportByVersion("Excel", 15, 16)]
        public Double Floor_Math(Double arg1, object arg2, object arg3)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor_Math", arg1, arg2, arg3);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double Floor_Math(Double arg1)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor_Math", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">Double arg1</param>
        /// <param name="arg2">optional object arg2</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public Double Floor_Math(Double arg1, object arg2)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "Floor_Math", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImSinh(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImSinh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">object arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public string ImCosh(object arg1)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "ImCosh", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        /// <param name="arg2">string arg2</param>
        [SupportByVersion("Excel", 15, 16)]
        public object FilterXML(string arg1, string arg2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FilterXML", arg1, arg2);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public object WebService(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "WebService", arg1);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="arg1">string arg1</param>
        [SupportByVersion("Excel", 15, 16)]
        public object EncodeURL(string arg1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "EncodeURL", arg1);
        }

        #endregion

        #pragma warning restore
    }
}

