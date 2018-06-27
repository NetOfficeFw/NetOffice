using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _CodePane
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _CodePane : COMObject, NetOffice.VBIDEApi._CodePane
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
                    _contractType = typeof(NetOffice.VBIDEApi._CodePane);
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
                    _type = typeof(_CodePane);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CodePane() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.CodePanes Collection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePanes>(this, "Collection", typeof(NetOffice.VBIDEApi.CodePanes));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.VBE VBE
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBE>(this, "VBE", typeof(NetOffice.VBIDEApi.VBE));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Window Window
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.Window>(this, "Window", typeof(NetOffice.VBIDEApi.Window));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int32 TopLine
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TopLine");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TopLine", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int32 CountOfVisibleLines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CountOfVisibleLines");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.CodeModule CodeModule
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodeModule>(this, "CodeModule", typeof(NetOffice.VBIDEApi.CodeModule));
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.Enums.vbext_CodePaneview CodePaneView
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.VBIDEApi.Enums.vbext_CodePaneview>(this, "CodePaneView");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void GetSelection(out Int32 startLine, out Int32 startColumn, out Int32 endLine, out Int32 endColumn)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true);
            startLine = 0;
            startColumn = 0;
            endLine = 0;
            endColumn = 0;
            object[] paramsArray = new object[] { startLine, startColumn, endLine, endColumn };

            InvokerService.InvokeInternal.ExecuteMethodExtended(this, "GetSelection", paramsArray, modifiers);

            startLine = (Int32)paramsArray[0];
            startColumn = (Int32)paramsArray[1];
            endLine = (Int32)paramsArray[2];
            endColumn = (Int32)paramsArray[3];
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void SetSelection(Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSelection", startLine, startColumn, endLine, endColumn);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void Show()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Show");
        }

        #endregion

        #pragma warning restore
    }
}
