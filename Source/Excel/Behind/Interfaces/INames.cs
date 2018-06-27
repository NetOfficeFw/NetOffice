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
    /// Interface INames 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    public class INames : COMObject, NetOffice.ExcelApi.INames
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
                    _contractType = typeof(NetOffice.ExcelApi.INames);
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
                    _type = typeof(INames);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public INames() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
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
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
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
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        /// <param name="nameLocal">optional object nameLocal</param>
        /// <param name="refersToLocal">optional object refersToLocal</param>
        /// <param name="categoryLocal">optional object categoryLocal</param>
        /// <param name="refersToR1C1">optional object refersToR1C1</param>
        /// <param name="refersToR1C1Local">optional object refersToR1C1Local</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1, refersToR1C1Local });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add()
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name));
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), name);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), name, refersTo);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), name, refersTo, visible);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), name, refersTo, visible, macroType);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        /// <param name="nameLocal">optional object nameLocal</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category, nameLocal });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        /// <param name="nameLocal">optional object nameLocal</param>
        /// <param name="refersToLocal">optional object refersToLocal</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        /// <param name="nameLocal">optional object nameLocal</param>
        /// <param name="refersToLocal">optional object refersToLocal</param>
        /// <param name="categoryLocal">optional object categoryLocal</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="name">optional object name</param>
        /// <param name="refersTo">optional object refersTo</param>
        /// <param name="visible">optional object visible</param>
        /// <param name="macroType">optional object macroType</param>
        /// <param name="shortcutKey">optional object shortcutKey</param>
        /// <param name="category">optional object category</param>
        /// <param name="nameLocal">optional object nameLocal</param>
        /// <param name="refersToLocal">optional object refersToLocal</param>
        /// <param name="categoryLocal">optional object categoryLocal</param>
        /// <param name="refersToR1C1">optional object refersToR1C1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "Add", typeof(NetOffice.ExcelApi.Name), new object[] { name, refersTo, visible, macroType, shortcutKey, category, nameLocal, refersToLocal, categoryLocal, refersToR1C1 });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Custom Indexer
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public virtual NetOffice.ExcelApi.Name this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", typeof(NetOffice.ExcelApi.Name), index);
            }
        }

        /// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Custom Indexer
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="indexLocal">optional object indexLocal</param>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
        public virtual NetOffice.ExcelApi.Name this[object index, object indexLocal]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", typeof(NetOffice.ExcelApi.Name), index, indexLocal);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="indexLocal">optional object indexLocal</param>
        /// <param name="refersTo">optional object refersTo</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.ExcelApi.Name this[object index, object indexLocal, object refersTo]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.Name>(this, "_Default", typeof(NetOffice.ExcelApi.Name), index, indexLocal, refersTo);
            }
        }

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.Name>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.Name>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.Name>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Name>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.Name> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.Name item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}

