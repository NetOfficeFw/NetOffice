using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface IApplicationEvents2 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class IApplicationEvents2 : COMObject, NetOffice.WordApi.IApplicationEvents2
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
                    _contractType = typeof(NetOffice.WordApi.IApplicationEvents2);
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
                    _type = typeof(IApplicationEvents2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IApplicationEvents2() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Startup()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Startup");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Quit()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Quit");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DocumentChange()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentChange");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DocumentOpen(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentOpen", doc);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DocumentBeforeClose(NetOffice.WordApi.Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforeClose", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DocumentBeforePrint(NetOffice.WordApi.Document doc, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforePrint", doc, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="saveAsUI">bool saveAsUI</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void DocumentBeforeSave(NetOffice.WordApi.Document doc, bool saveAsUI, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "DocumentBeforeSave", doc, saveAsUI, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void NewDocument(NetOffice.WordApi.Document doc)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "NewDocument", doc);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WindowActivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowActivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="doc">NetOffice.WordApi.Document doc</param>
		/// <param name="wn">NetOffice.WordApi.Window wn</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WindowDeactivate(NetOffice.WordApi.Document doc, NetOffice.WordApi.Window wn)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowDeactivate", doc, wn);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WindowSelectionChange(NetOffice.WordApi.Selection sel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowSelectionChange", sel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WindowBeforeRightClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowBeforeRightClick", sel, cancel);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sel">NetOffice.WordApi.Selection sel</param>
		/// <param name="cancel">bool cancel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void WindowBeforeDoubleClick(NetOffice.WordApi.Selection sel, bool cancel)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "WindowBeforeDoubleClick", sel, cancel);
		}

		#endregion

		#pragma warning restore
	}
}

