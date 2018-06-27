using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface Crop 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860761.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Crop : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.Crop
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
                    _contractType = typeof(NetOffice.OfficeApi.Crop);
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
                    _type = typeof(Crop);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Crop() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862450.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single PictureOffsetX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PictureOffsetX");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureOffsetX", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864637.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single PictureOffsetY
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PictureOffsetY");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureOffsetY", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860544.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single PictureWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PictureWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860512.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single PictureHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "PictureHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureHeight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861232.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single ShapeLeft
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ShapeLeft");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShapeLeft", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861517.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single ShapeTop
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ShapeTop");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShapeTop", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861716.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single ShapeWidth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ShapeWidth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShapeWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864643.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Single ShapeHeight
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "ShapeHeight");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShapeHeight", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
