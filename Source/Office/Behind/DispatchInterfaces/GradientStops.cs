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
    /// DispatchInterface GradientStops 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861159.aspx </remarks>
    public class GradientStops : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.GradientStops
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
                    _contractType = typeof(NetOffice.OfficeApi.GradientStops);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public GradientStops() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">Int32 index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.OfficeApi.GradientStop this[Int32 index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GradientStop>(this, "Item", typeof(NetOffice.OfficeApi.GradientStop), index);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864855.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx </remarks>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete(object index)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861233.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Insert(Int32 rGB, Single position, object transparency, object index)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", rGB, position, transparency, index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Insert(Int32 rGB, Single position)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", rGB, position);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863159.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Insert(Int32 rGB, Single position, object transparency)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert", rGB, position, transparency);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        /// <param name="brightness">optional Single Brightness = 0</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Insert2(Int32 rGB, Single position, object transparency, object index, object brightness)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2", new object[] { rGB, position, transparency, index, brightness });
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Insert2(Int32 rGB, Single position)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2", rGB, position);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Insert2(Int32 rGB, Single position, object transparency)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2", rGB, position, transparency);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864086.aspx </remarks>
        /// <param name="rGB">Int32 rGB</param>
        /// <param name="position">Single position</param>
        /// <param name="transparency">optional Single Transparency = 0</param>
        /// <param name="index">optional Int32 Index = -1</param>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual void Insert2(Int32 rGB, Single position, object transparency, object index)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Insert2", rGB, position, transparency, index);
        }

        #endregion

        #region IEnumerableProvider<NetOffice.OfficeApi.GradientStop>

        ICOMObject IEnumerableProvider<NetOffice.OfficeApi.GradientStop>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.OfficeApi.GradientStop> GetEnumerator()
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}
