using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// DispatchInterface PageFields 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class PageFields : COMObject, IEnumerableProvider<NetOffice.OWC10Api.PageField>
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
                    _type = typeof(PageFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PageFields(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PageFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageFields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OWC10Api.PageField this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageField>(this, "Item", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
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
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("OWC10", 1)]
		public void Delete(object index)
		{
			 Factory.ExecuteMethod(this, "Delete", index);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object index</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "Add", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, new object[]{ source, fieldType, name, totalType, index });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "Add", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "Add", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "Add", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField Add(object source, object fieldType, object name, object totalType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "Add", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType, name, totalType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		/// <param name="index">optional object index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType, object index)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "AddBroken", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, new object[]{ source, fieldType, name, totalType, index });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "AddBroken", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "AddBroken", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "AddBroken", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType, name);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="source">object source</param>
		/// <param name="fieldType">optional object fieldType</param>
		/// <param name="name">optional object name</param>
		/// <param name="totalType">optional NetOffice.OWC10Api.Enums.DscTotalTypeEnum TotalType = 0</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.PageField AddBroken(object source, object fieldType, object name, object totalType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.PageField>(this, "AddBroken", NetOffice.OWC10Api.PageField.LateBindingApiWrapperType, source, fieldType, name, totalType);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OWC10Api.PageField>

        ICOMObject IEnumerableProvider<NetOffice.OWC10Api.PageField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OWC10Api.PageField>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OWC10Api.PageField>

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public IEnumerator<NetOffice.OWC10Api.PageField> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OWC10Api.PageField item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}