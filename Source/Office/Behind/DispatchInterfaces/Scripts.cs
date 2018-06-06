using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface Scripts 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
    public class Scripts : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.Scripts
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Instance Type
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
                    _type = typeof(Scripts);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Scripts() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.Script this[object index]
        {
            get
            {
                return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Item", typeof(NetOffice.OfficeApi.Script), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        /// <param name="extended">optional string Extended = </param>
        /// <param name="scriptText">optional string ScriptText = </param>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended, object scriptText)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), new object[] { anchor, location, language, id, extended, scriptText });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add()
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script));
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), anchor);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor, object location)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), anchor, location);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor, object location, object language)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), anchor, location, language);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), anchor, location, language, id);
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
        /// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
        /// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
        /// <param name="id">optional string Id = </param>
        /// <param name="extended">optional string Extended = </param>
        [CustomMethod]
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended)
        {
            return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", typeof(NetOffice.OfficeApi.Script), new object[] { anchor, location, language, id, extended });
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.Script>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.Script>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.Script>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.Script>

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.OfficeApi.Script> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.Script item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
