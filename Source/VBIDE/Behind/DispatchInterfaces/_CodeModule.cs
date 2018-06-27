using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VBIDEApi.Behind
{
    /// <summary>
    /// DispatchInterface _CodeModule
    /// SupportByVersion VBIDE, 12,14,5.3
    /// </summary>
    [SupportByVersion("VBIDE", 12, 14, 5.3)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class _CodeModule : COMObject, NetOffice.VBIDEApi._CodeModule
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
                    _contractType = typeof(NetOffice.VBIDEApi._CodeModule);
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
                    _type = typeof(_CodeModule);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CodeModule() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.VBComponent Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.VBComponent>(this, "Parent", typeof(NetOffice.VBIDEApi.VBComponent));
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
        /// Get/Set
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="count">Int32 count</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_Lines(Int32 startLine, Int32 count)
        {
            return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Lines", startLine, count);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_Lines
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="count">Int32 count</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_Lines")]
        public virtual string Lines(Int32 startLine, Int32 count)
        {
            return get_Lines(startLine, count);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int32 CountOfLines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CountOfLines");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcStartLine", procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcStartLine
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcStartLine")]
        public virtual Int32 ProcStartLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return get_ProcStartLine(procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcCountLines", procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcCountLines
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcCountLines")]
        public virtual Int32 ProcCountLines(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return get_ProcCountLines(procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 get_ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ProcBodyLine", procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcBodyLine
        /// </summary>
        /// <param name="procName">string procName</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcBodyLine")]
        public virtual Int32 ProcBodyLine(string procName, NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return get_ProcBodyLine(procName, procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string get_ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
            procKind = 0;
            object[] paramsArray = new object[] { line, procKind };

            
            string returnItem = InvokerService.InvokeInternal.ExecuteStringPropertyGetExtended(this, "ProcOfLine", paramsArray, modifiers);

            procKind = (NetOffice.VBIDEApi.Enums.vbext_ProcKind)paramsArray[1];
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Alias for get_ProcOfLine
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="procKind">NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3), Redirect("get_ProcOfLine")]
        public virtual string ProcOfLine(Int32 line, out NetOffice.VBIDEApi.Enums.vbext_ProcKind procKind)
        {
            return get_ProcOfLine(line, out procKind);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int32 CountOfDeclarationLines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "CountOfDeclarationLines");
            }
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// Get
        /// </summary>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual NetOffice.VBIDEApi.CodePane CodePane
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.VBIDEApi.CodePane>(this, "CodePane", typeof(NetOffice.VBIDEApi.CodePane));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="_string">string string</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void AddFromString(string _string)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddFromString", _string);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void AddFromFile(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AddFromFile", fileName);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="_string">string string</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void InsertLines(Int32 line, string _string)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "InsertLines", line, _string);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="count">optional Int32 Count = 1</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void DeleteLines(Int32 startLine, object count)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteLines", startLine, count);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="startLine">Int32 startLine</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void DeleteLines(Int32 startLine)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteLines", startLine);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="line">Int32 line</param>
        /// <param name="_string">string string</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual void ReplaceLine(Int32 line, string _string)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ReplaceLine", line, _string);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="eventName">string eventName</param>
        /// <param name="objectName">string objectName</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual Int32 CreateEventProc(string eventName, string objectName)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CreateEventProc", eventName, objectName);
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        /// <param name="matchCase">optional bool MatchCase = false</param>
        /// <param name="patternSearch">optional bool PatternSearch = false</param>
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase, object patternSearch)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[] { target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase, patternSearch });
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[] { target, startLine, startColumn, endLine, endColumn });
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[] { target, startLine, startColumn, endLine, endColumn, wholeWord });
        }

        /// <summary>
        /// SupportByVersion VBIDE 12, 14, 5.3
        /// </summary>
        /// <param name="target">string target</param>
        /// <param name="startLine">Int32 startLine</param>
        /// <param name="startColumn">Int32 startColumn</param>
        /// <param name="endLine">Int32 endLine</param>
        /// <param name="endColumn">Int32 endColumn</param>
        /// <param name="wholeWord">optional bool WholeWord = false</param>
        /// <param name="matchCase">optional bool MatchCase = false</param>
        [CustomMethod]
        [SupportByVersion("VBIDE", 12, 14, 5.3)]
        public virtual bool Find(string target, Int32 startLine, Int32 startColumn, Int32 endLine, Int32 endColumn, object wholeWord, object matchCase)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Find", new object[] { target, startLine, startColumn, endLine, endColumn, wholeWord, matchCase });
        }

        #endregion

        #pragma warning restore
    }
}
