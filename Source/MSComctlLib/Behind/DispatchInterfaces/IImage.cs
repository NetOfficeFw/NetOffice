using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSComctlLibApi;

namespace NetOffice.MSComctlLibApi.Behind
{
	/// <summary>
	/// DispatchInterface IImage 
	/// SupportByVersion MSComctlLib, 6
	/// </summary>
	[SupportByVersion("MSComctlLib", 6)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IImage : COMObject, NetOffice.MSComctlLibApi.IImage
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
                    _contractType = typeof(NetOffice.MSComctlLibApi.IImage);
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
                    _type = typeof(IImage);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IImage() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual Int16 Index
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "Index");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Index", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual string Key
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Key");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Key", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual object Tag
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Tag");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Tag", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Picture Picture
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Picture", paramsArray);
                return returnItem as stdole.Picture;
            }
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Picture", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		/// <param name="style">optional object style</param>
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Draw(Int32 hDC, object x, object y, object style)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Draw", hDC, x, y, style);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Draw(Int32 hDC)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Draw", hDC);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Draw(Int32 hDC, object x)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Draw", hDC, x);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		/// <param name="hDC">Int32 hDC</param>
		/// <param name="x">optional object x</param>
		/// <param name="y">optional object y</param>
		[CustomMethod]
		[SupportByVersion("MSComctlLib", 6)]
		public virtual void Draw(Int32 hDC, object x, object y)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Draw", hDC, x, y);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// </summary>
		[SupportByVersion("MSComctlLib", 6), NativeResult]
		public virtual stdole.Picture ExtractIcon()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "ExtractIcon", paramsArray);
            return returnItem as stdole.Picture;
        }

		#endregion

		#pragma warning restore
	}
}

