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
	/// DispatchInterface GradientStops 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "Item")]
	public class GradientStops : _IMsoDispObj, IEnumerableProvider<NetOffice.OfficeApi.GradientStop>
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
                    _type = typeof(GradientStops);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public GradientStops(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public GradientStops(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public GradientStops(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Office", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.OfficeApi.GradientStop this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GradientStop>(this, "Item", NetOffice.OfficeApi.GradientStop.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Count"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
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
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Delete"/> </remarks>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Delete(object index)
		{
			 Factory.ExecuteMethod(this, "Delete", index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Delete"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[SupportByVersion("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position, object transparency, object index)
		{
			 Factory.ExecuteMethod(this, "Insert", rGB, position, transparency, index);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position)
		{
			 Factory.ExecuteMethod(this, "Insert", rGB, position);
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 12,14,15,16)]
		public void Insert(Int32 rGB, Single position, object transparency)
		{
			 Factory.ExecuteMethod(this, "Insert", rGB, position, transparency);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert2"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		/// <param name="brightness">optional Single Brightness = 0</param>
		[SupportByVersion("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency, object index, object brightness)
		{
			 Factory.ExecuteMethod(this, "Insert2", new object[]{ rGB, position, transparency, index, brightness });
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert2"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position)
		{
			 Factory.ExecuteMethod(this, "Insert2", rGB, position);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert2"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency)
		{
			 Factory.ExecuteMethod(this, "Insert2", rGB, position, transparency);
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.GradientStops.Insert2"/> </remarks>
		/// <param name="rGB">Int32 rGB</param>
		/// <param name="position">Single position</param>
		/// <param name="transparency">optional Single Transparency = 0</param>
		/// <param name="index">optional Int32 Index = -1</param>
		[CustomMethod]
		[SupportByVersion("Office", 14,15,16)]
		public void Insert2(Int32 rGB, Single position, object transparency, object index)
		{
			 Factory.ExecuteMethod(this, "Insert2", rGB, position, transparency, index);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.GradientStop>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.GradientStop>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.OfficeApi.GradientStop>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.OfficeApi.GradientStop>

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.OfficeApi.GradientStop> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.OfficeApi.GradientStop item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Office, 12,14,15,16
        /// </summary>
        [SupportByVersion("Office", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}