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
	/// DispatchInterface PictureEffects 
	/// SupportByVersion Office, 14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects"/> </remarks>
	[SupportByVersion("Office", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class PictureEffects : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>
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
                    _type = typeof(PictureEffects);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PictureEffects(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PictureEffects(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PictureEffects(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.PictureEffect this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.PictureEffect>(this, "Item", NetOffice.OfficeApi.PictureEffect.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects.Count"/> </remarks>
		[SupportByVersion("Office", 14,15,16)]
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
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects.Insert"/> </remarks>
		/// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
		/// <param name="position">optional Int32 Position = -1</param>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType, object position)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PictureEffect>(this, "Insert", NetOffice.OfficeApi.PictureEffect.LateBindingApiWrapperType, effectType, position);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects.Insert"/> </remarks>
		/// <param name="effectType">NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.PictureEffect Insert(NetOffice.OfficeApi.Enums.MsoPictureEffectType effectType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OfficeApi.PictureEffect>(this, "Insert", NetOffice.OfficeApi.PictureEffect.LateBindingApiWrapperType, effectType);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects.Delete"/> </remarks>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("Office", 14,15,16)]
		public void Delete(object index)
		{
			 Factory.ExecuteMethod(this, "Delete", index);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.PictureEffects.Delete"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.PictureEffect>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.PictureEffect>

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.PictureEffect> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.PictureEffect item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 14,15,16
        /// </summary>
        [SupportByVersion("Office", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}