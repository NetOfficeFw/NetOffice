using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.VBIDEApi
{
	/// <summary>
	/// DispatchInterface _Components 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class _Components : COMObject, IEnumerableProvider<NetOffice.VBIDEApi.Component>
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
                    _type = typeof(_Components);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Components(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Components(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Components(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[BaseResult]
		public NetOffice.VBIDEApi.Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.VBIDEApi.Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBProject Parent
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBProject>(this, "Parent", NetOffice.VBIDEApi.VBProject.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// Get
		/// </summary>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBE VBE
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", NetOffice.VBIDEApi.VBE.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.VBIDEApi.Component this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.Component>(this, "Item", NetOffice.VBIDEApi.Component.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="component">NetOffice.VBIDEApi.Component component</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void Remove(NetOffice.VBIDEApi.Component component)
		{
			 Factory.ExecuteMethod(this, "Remove", component);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="componentType">NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Component Add(NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.Component>(this, "Add", NetOffice.VBIDEApi.Component.LateBindingApiWrapperType, componentType);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.Component Import(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.Component>(this, "Import", NetOffice.VBIDEApi.Component.LateBindingApiWrapperType, fileName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VBIDEApi.Component>

        ICOMObject IEnumerableProvider<NetOffice.VBIDEApi.Component>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.VBIDEApi.Component>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.Component>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public IEnumerator<NetOffice.VBIDEApi.Component> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VBIDEApi.Component item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12,14,5.3)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion

		#pragma warning restore
	}
}