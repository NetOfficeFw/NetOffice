using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface ICalculatedMembers 
    /// SupportByVersion Excel, 10,11,12,14,15,16
    /// </summary>
    public class ICalculatedMembers : COMObject, NetOffice.ExcelApi.ICalculatedMembers
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
                    _contractType = typeof(NetOffice.ExcelApi.ICalculatedMembers);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type        /// </summary>
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
                    _type = typeof(ICalculatedMembers);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICalculatedMembers() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.ExcelApi.CalculatedMember this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Default", typeof(NetOffice.ExcelApi.CalculatedMember), index);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder, object type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder, type);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="hierarchizeDistinct">optional object hierarchizeDistinct</param>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder, object hierarchizeDistinct)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, dynamic, displayFolder, hierarchizeDistinct });
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, string formula)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula);
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, string formula, object solveOrder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, dynamic });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="dynamic">optional object dynamic</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember Add(string name, object formula, object solveOrder, object type, object dynamic, object displayFolder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "Add", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, dynamic, displayFolder });
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder, object type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder, type);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">string formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember _Add(string name, string formula, object solveOrder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "_Add", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        /// <param name="parentMember">optional object parentMember</param>
        /// <param name="numberFormat">optional object numberFormat</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember, object numberFormat)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, parentMember, numberFormat });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), name, formula, solveOrder, type);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, displayFolder });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, displayFolder, measureGroup });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy });
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="name">string name</param>
        /// <param name="formula">object formula</param>
        /// <param name="solveOrder">optional object solveOrder</param>
        /// <param name="type">optional object type</param>
        /// <param name="displayFolder">optional object displayFolder</param>
        /// <param name="measureGroup">optional object measureGroup</param>
        /// <param name="parentHierarchy">optional object parentHierarchy</param>
        /// <param name="parentMember">optional object parentMember</param>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual NetOffice.ExcelApi.CalculatedMember AddCalculatedMember(string name, object formula, object solveOrder, object type, object displayFolder, object measureGroup, object parentHierarchy, object parentMember)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.CalculatedMember>(this, "AddCalculatedMember", typeof(NetOffice.ExcelApi.CalculatedMember), new object[] { name, formula, solveOrder, type, displayFolder, measureGroup, parentHierarchy, parentMember });
        }

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.CalculatedMember>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.CalculatedMember>

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.CalculatedMember> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.CalculatedMember item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}

