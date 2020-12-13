using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Scripts 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Scripts : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.Script>
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Scripts(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Scripts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Scripts(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public Int32 Count
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.Script this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Item", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, index);
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended, object scriptText)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, new object[]{ anchor, location, language, id, extended, scriptText });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, anchor);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, anchor, location);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, anchor, location, language);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="anchor">optional object Anchor = null (Nothing in visual basic)</param>
		/// <param name="location">optional NetOffice.OfficeApi.Enums.MsoScriptLocation Location = 2</param>
		/// <param name="language">optional NetOffice.OfficeApi.Enums.MsoScriptLanguage Language = 2</param>
		/// <param name="id">optional string Id = </param>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, anchor, location, language, id);
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
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Script Add(object anchor, object location, object language, object id, object extended)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.Script>(this, "Add", NetOffice.OfficeApi.Script.LateBindingApiWrapperType, new object[]{ anchor, location, language, id, extended });
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.Script>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.Script>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
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
        public IEnumerator<NetOffice.OfficeApi.Script> GetEnumerator()
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
        [SupportByVersion("Office", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}