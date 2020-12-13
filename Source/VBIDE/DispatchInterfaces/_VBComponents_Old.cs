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
	/// DispatchInterface _VBComponents_Old 
	/// SupportByVersion VBIDE, 12,14,5.3
	/// </summary>
	[SupportByVersion("VBIDE", 12,14,5.3)]
	[EntityType(EntityType.IsDispatchInterface), BaseType, Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class _VBComponents_Old : COMObject, IEnumerableProvider<NetOffice.VBIDEApi.VBComponent>
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
                    _type = typeof(_VBComponents_Old);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _VBComponents_Old(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _VBComponents_Old(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _VBComponents_Old(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

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
		public NetOffice.VBIDEApi.VBComponent this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.VBComponent>(this, "Item", NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="vBComponent">NetOffice.VBIDEApi.VBComponent vBComponent</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public void Remove(NetOffice.VBIDEApi.VBComponent vBComponent)
		{
			 Factory.ExecuteMethod(this, "Remove", vBComponent);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="componentType">NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBComponent Add(NetOffice.VBIDEApi.Enums.vbext_ComponentType componentType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.VBComponent>(this, "Add", NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType, componentType);
		}

		/// <summary>
		/// SupportByVersion VBIDE 12, 14, 5.3
		/// </summary>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("VBIDE", 12,14,5.3)]
		public NetOffice.VBIDEApi.VBComponent Import(string fileName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.VBIDEApi.VBComponent>(this, "Import", NetOffice.VBIDEApi.VBComponent.LateBindingApiWrapperType, fileName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.VBIDEApi.VBComponent>

        ICOMObject IEnumerableProvider<NetOffice.VBIDEApi.VBComponent>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.VBIDEApi.VBComponent>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.VBIDEApi.VBComponent>

        /// <summary>
        /// SupportByVersion VBIDE, 12,14,5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public IEnumerator<NetOffice.VBIDEApi.VBComponent> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.VBIDEApi.VBComponent item in innerEnumerator)
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