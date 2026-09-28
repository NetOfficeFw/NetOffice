using System;
using System.Collections.Generic;
using System.Linq;

namespace NetOffice.CodeGen.Emit
{
    internal static class InvocationBodyRenderer
    {
        public static string Render(WrapperInvocation invocation)
        {
            if (invocation == null) throw new ArgumentNullException(nameof(invocation));
            invocation.Validate();
            if (invocation.Kind == WrapperInvocationKind.EventRaise) return RenderEventRaise(invocation);
            if (invocation.Kind == WrapperInvocationKind.LocalForward) return RenderLocalForward(invocation);
            return invocation.Api == WrapperInvocationApi.Factory ? RenderFactory(invocation) : RenderInvoker(invocation);
        }
        public static string Render(WrapperInvocation invocation, WrapperModuleRuntime moduleRuntime)
        {
            if (moduleRuntime == null) return Render(invocation);
            if (invocation == null) throw new ArgumentNullException(nameof(invocation));
            if (invocation.Kind == WrapperInvocationKind.LocalForward) return Render(invocation);
            return Render(new WrapperInvocation
            {
                Api = invocation.Api,
                Kind = invocation.Kind,
                ReturnKind = invocation.ReturnKind,
                Target = string.IsNullOrWhiteSpace(invocation.Target) || string.Equals(invocation.Target, "this", StringComparison.Ordinal) ? moduleRuntime.InstanceField : invocation.Target,
                DispatchName = invocation.DispatchName,
                ReturnType = invocation.ReturnType,
                WrapperTypeExpression = invocation.WrapperTypeExpression,
                ScalarConversion = invocation.ScalarConversion,
                EventName = invocation.EventName,
                ArgumentPacking = invocation.ArgumentPacking,
                ObjectArrayStyle = invocation.ObjectArrayStyle,
                RawCallCast = invocation.RawCallCast,
                KnownReferenceFactoryStyle = invocation.KnownReferenceFactoryStyle,
                ReturnValueStyle = invocation.ReturnValueStyle,
                ReturnLocalName = invocation.ReturnLocalName,
                ReturnLocalType = invocation.ReturnLocalType,
                ReturnCastStyle = invocation.ReturnCastStyle,
                InvokerCallStyle = invocation.InvokerCallStyle,
                StatementTerminatorStyle = invocation.StatementTerminatorStyle,
                FactoryExpression = moduleRuntime.FactoryProperty,
                InvokerExpression = moduleRuntime.InvokerProperty,
                ReleaseArguments = invocation.ReleaseArguments,
                Arguments = invocation.Arguments
            });
        }

        private static string RenderLocalForward(WrapperInvocation value)
        {
            var target = string.IsNullOrWhiteSpace(value.Target) ? string.Empty : value.Target + ".";
            var call = target + value.DispatchName + "(" + string.Join(", ", value.Arguments.Select(argument => argument.Expression)) + ")";
            return value.ReturnKind == WrapperReturnKind.Void ? call + ";" : "return " + call + ";";
        }


        private static string RenderFactory(WrapperInvocation value)
        {
            var arguments = value.Arguments.Select(x => x.Expression).ToList();
            var targetAndName = value.Target + ", " + Literal(value.DispatchName);
            var maxArguments = value.Kind == WrapperInvocationKind.Method ? 8 : 4;
            string call;
            switch (value.Kind)
            {
                case WrapperInvocationKind.PropertySet:
                case WrapperInvocationKind.PropertySetVariant:
                case WrapperInvocationKind.PropertySetValue:
                case WrapperInvocationKind.PropertySetEnum:
                case WrapperInvocationKind.PropertyPutRef:
                    var propertyValueIndex = -1;
                    for (var index = 0; index < value.Arguments.Count; index++)
                    {
                        if (value.Arguments[index].IsPropertyValue)
                        {
                            propertyValueIndex = index;
                            break;
                        }
                    }
                    if (propertyValueIndex < 0) propertyValueIndex = arguments.Count - 1;
                    var newValue = arguments[propertyValueIndex];
                    var indexArguments = new List<string>(arguments.Count - 1);
                    for (var index = 0; index < arguments.Count; index++)
                        if (index != propertyValueIndex) indexArguments.Add(arguments[index]);
                    if (value.ArgumentPacking == WrapperInvocationArgumentPacking.ObjectArray || indexArguments.Count > maxArguments) indexArguments = new List<string> { ArrayExpression(indexArguments, value.ObjectArrayStyle) };
                    var setter = value.Kind == WrapperInvocationKind.PropertyPutRef ? "ExecuteReferencePropertySet" : value.Kind == WrapperInvocationKind.PropertySetVariant ? "ExecuteVariantPropertySet" : value.Kind == WrapperInvocationKind.PropertySetEnum ? "ExecuteEnumPropertySet" : value.Kind == WrapperInvocationKind.PropertySetValue ? "ExecuteValuePropertySet" : "ExecutePropertySet";
                    call = value.FactoryExpression + "." + setter + "(" + JoinCall(targetAndName, new[] { newValue }.Concat(indexArguments)) + ")";
                    return call + ";";
                case WrapperInvocationKind.PropertyGet:
                    if (value.ArgumentPacking == WrapperInvocationArgumentPacking.ObjectArray || arguments.Count > maxArguments) arguments = new List<string> { ArrayExpression(arguments, value.ObjectArrayStyle) };
                    call = FactoryGetCall(value, "PropertyGet", targetAndName, arguments);
                    break;
                default:
                    if (value.ArgumentPacking == WrapperInvocationArgumentPacking.ObjectArray || arguments.Count > maxArguments) arguments = new List<string> { ArrayExpression(arguments, value.ObjectArrayStyle) };
                    if (value.ReturnKind == WrapperReturnKind.Void)
                        return value.FactoryExpression + ".ExecuteMethod(" + JoinCall(targetAndName, arguments) + ");";
                    call = FactoryGetCall(value, "MethodGet", targetAndName, arguments);
                    break;
            }
            return "return " + call + ";";
        }

        private static string FactoryGetCall(WrapperInvocation value, string suffix, string targetAndName, IList<string> arguments)
        {
            switch (value.ReturnKind)
            {
                case WrapperReturnKind.KnownReference:
                    return value.FactoryExpression + ".ExecuteKnownReference" + suffix + "<" + value.ReturnType + ">(" + JoinCall(targetAndName, new[] { value.WrapperTypeExpression }.Concat(arguments)) + ")";
                case WrapperReturnKind.BaseReference:
                    return value.FactoryExpression + ".ExecuteBaseReference" + suffix + "<" + value.ReturnType + ">(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Enum:
                    return value.FactoryExpression + ".ExecuteEnum" + suffix + "<" + value.ReturnType + ">(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Struct:
                    return value.FactoryExpression + ".ExecuteStruct" + suffix + "<" + value.ReturnType + ">(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Variant:
                    return value.FactoryExpression + ".ExecuteVariant" + suffix + "(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Reference:
                    return value.FactoryExpression + ".ExecuteReference" + suffix + "<" + value.ReturnType + ">(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.UntypedReference:
                    return value.FactoryExpression + ".ExecuteReference" + suffix + "(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Raw:
                    return "(" + value.ReturnType + ")" + value.FactoryExpression + ".ExecuteObject" + suffix + "(" + JoinCall(targetAndName, arguments) + ")";
                case WrapperReturnKind.Scalar:
                    return value.FactoryExpression + ".Execute" + FactoryScalarName(value, suffix) + suffix + "(" + JoinCall(targetAndName, arguments) + ")";
                default:
                    throw new InvalidOperationException("Unsupported Factory return kind: " + value.ReturnKind);
            }
        }

        private static string RenderInvoker(WrapperInvocation value)
        {
            if (value.InvokerCallStyle == WrapperInvokerCallStyle.Direct) return RenderDirectInvoker(value);
            var lines = new List<string>();
            var arguments = value.Arguments.Select(x => x.Expression).ToList();
            var hasByRefArguments = value.Arguments.Any(argument => argument.ByRef);
            if (hasByRefArguments)
                lines.Add("ParameterModifier[] modifiers = " + value.InvokerExpression + ".CreateParamModifiers(" + string.Join(", ", value.Arguments.Select(argument => argument.ByRef ? "true" : "false")) + ");");
            foreach (var argument in value.Arguments.Where(argument => !string.IsNullOrWhiteSpace(argument.InitializationExpression)))
                lines.Add((string.IsNullOrWhiteSpace(argument.WriteBackExpression) ? argument.Expression : argument.WriteBackExpression) + " = " + argument.InitializationExpression + ";");
            lines.Add("object[] paramsArray = " + (arguments.Count == 0 ? "null" : value.InvokerExpression + ".ValidateParamsArray(" + string.Join(", ", arguments) + ")") + ";");
            var modifiers = hasByRefArguments ? ", modifiers" : string.Empty;
            if (value.ReleaseArguments) lines.Add("try");
            if (value.ReleaseArguments) lines.Add("{");
            var indent = value.ReleaseArguments ? "    " : string.Empty;
            string call;
            switch (value.Kind)
            {
                case WrapperInvocationKind.PropertySet:
                case WrapperInvocationKind.PropertySetVariant:
                case WrapperInvocationKind.PropertySetValue:
                case WrapperInvocationKind.PropertySetEnum:
                case WrapperInvocationKind.PropertyPutRef:
                    call = value.InvokerExpression + ".PropertySet(" + value.Target + ", " + Literal(value.DispatchName) + ", paramsArray" + modifiers + ");";
                    lines.Add(indent + Terminate(call, value.StatementTerminatorStyle));
                    AddWriteBack(lines, value, indent);
                    break;
                case WrapperInvocationKind.PropertyGet:
                    call = value.InvokerExpression + ".PropertyGet(" + value.Target + ", " + Literal(value.DispatchName) + ", paramsArray" + modifiers + ")";
                    AddReturn(lines, value, call, indent);
                    break;
                default:
                    if (value.ReturnKind == WrapperReturnKind.Void)
                    {
                        lines.Add(indent + Terminate(value.InvokerExpression + ".Method(" + value.Target + ", " + Literal(value.DispatchName) + ", paramsArray" + modifiers + ");", value.StatementTerminatorStyle));
                        AddWriteBack(lines, value, indent);
                    }
                    else
                    {
                        call = value.InvokerExpression + ".MethodReturn(" + value.Target + ", " + Literal(value.DispatchName) + ", paramsArray" + modifiers + ")";
                        AddReturn(lines, value, call, indent);
                    }
                    break;
            }
            if (value.ReleaseArguments)
            {
                lines.Add("}");
                lines.Add("finally");
                lines.Add("{");
                lines.Add("    if (paramsArray != null) " + value.InvokerExpression + ".ReleaseParamsArray(paramsArray);");
                lines.Add("}");
            }
            return string.Join("\n", lines);
        }

        private static string RenderDirectInvoker(WrapperInvocation value)
        {
            var operation = value.Kind == WrapperInvocationKind.PropertyGet ? "PropertyGet" : "Method";
            var call = value.InvokerExpression + "." + operation + "(" + value.Target + ", " + Literal(value.DispatchName) + ")";
            if (value.ReturnKind == WrapperReturnKind.Void)
                return Terminate(call + ";", value.StatementTerminatorStyle);
            var cast = value.RawCallCast == WrapperRawCallCast.Object ? "(object)" : string.Empty;
            return "return " + cast + call + ";";
        }

        private static string Terminate(string statement, WrapperStatementTerminatorStyle style)
        {
            return style == WrapperStatementTerminatorStyle.DoubleSemicolon ? statement + ";" : statement;
        }

        private static void AddReturn(List<string> lines, WrapperInvocation value, string call, string indent)
        {
            lines.Add(indent + "object returnItem = " + (value.ReturnKind == WrapperReturnKind.Raw && value.RawCallCast == WrapperRawCallCast.Object ? "(object)" : string.Empty) + call + ";");
            AddWriteBack(lines, value, indent);
            var referenceType = value.ReturnValueStyle == WrapperReturnValueStyle.Local && !string.IsNullOrWhiteSpace(value.ReturnLocalType) ? value.ReturnLocalType : value.ReturnType;
            switch (value.ReturnKind)
            {
                case WrapperReturnKind.Raw:
                    lines.Add(indent + "return " + (value.ReturnType == "object" ? "returnItem" : "(" + value.ReturnType + ")returnItem") + ";");
                    break;
                case WrapperReturnKind.Scalar:
                    lines.Add(indent + "return " + ScalarConversion(value, "returnItem") + ";");
                    break;
                case WrapperReturnKind.Enum:
                    lines.Add(indent + "int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);");
                    lines.Add(indent + "return (" + value.ReturnType + ")intReturnItem;");
                    break;
                case WrapperReturnKind.Struct:
                    lines.Add(indent + "return (" + value.ReturnType + ")returnItem;");
                    break;
                case WrapperReturnKind.Variant:
                    lines.Add(indent + "return returnItem;");
                    break;
                case WrapperReturnKind.Reference:
                case WrapperReturnKind.UntypedReference:
                case WrapperReturnKind.BaseReference:
                    AddReferenceReturn(lines, value, ApplyReferenceCast(value.FactoryExpression + ".CreateObjectFromComProxy(" + value.Target + ", returnItem)", referenceType, value.ReturnCastStyle), indent);
                    break;
                case WrapperReturnKind.KnownReference:
                    var knownCall = value.KnownReferenceFactoryStyle == WrapperKnownReferenceFactoryStyle.NonGeneric
                        ? ApplyReferenceCast(value.FactoryExpression + ".CreateKnownObjectFromComProxy(" + value.Target + ", returnItem, " + value.WrapperTypeExpression + ")", referenceType, value.ReturnCastStyle)
                        : ApplyReferenceCast(value.FactoryExpression + ".CreateKnownObjectFromComProxy<" + referenceType + ">(" + value.Target + ", returnItem, " + value.WrapperTypeExpression + ")", referenceType, value.ReturnCastStyle);
                    AddReferenceReturn(lines, value, knownCall, indent);
                    break;
                case WrapperReturnKind.EventArgument:
                    AddReferenceReturn(lines, value, ApplyReferenceCast(value.FactoryExpression + ".CreateEventArgumentObjectFromComProxy(" + value.Target + ", returnItem)", referenceType, value.ReturnCastStyle), indent);
                    break;
                default:
                    throw new InvalidOperationException("Unsupported Invoker return kind: " + value.ReturnKind);
            }
        }

        private static string ApplyReferenceCast(string expression, string returnType, WrapperReturnCastStyle style)
        {
            if (style == WrapperReturnCastStyle.Default) return expression;
            return style == WrapperReturnCastStyle.Explicit ? "(" + returnType + ")" + expression : expression + " as " + returnType;
        }

        private static void AddReferenceReturn(ICollection<string> lines, WrapperInvocation value, string expression, string indent)
        {
            if (value.ReturnValueStyle == WrapperReturnValueStyle.Local)
            {
                lines.Add(indent + (string.IsNullOrWhiteSpace(value.ReturnLocalType) ? value.ReturnType : value.ReturnLocalType) + " " + value.ReturnLocalName + " = " + expression + ";");
                lines.Add(indent + "return " + value.ReturnLocalName + ";");
            }
            else
            {
                lines.Add(indent + "return " + expression + ";");
            }
        }

        private static void AddWriteBack(ICollection<string> lines, WrapperInvocation value, string indent)
        {
            for (var index = 0; index < value.Arguments.Count; index++)
            {
                var argument = value.Arguments[index];
                if (!string.IsNullOrWhiteSpace(argument.WriteBackExpression))
                {
                    var source = "paramsArray[" + index + "]";
                    var converted = string.IsNullOrWhiteSpace(argument.WriteBackConversion)
                        ? "(" + argument.WriteBackType + ")" + source
                        : argument.WriteBackConversion.Replace("{0}", source);
                    lines.Add(indent + argument.WriteBackExpression + " = " + converted + ";");
                }
            }
        }

        private static string RenderEventRaise(WrapperInvocation value)
        {
            var eventName = string.IsNullOrWhiteSpace(value.EventName) ? value.DispatchName : value.EventName;
            var expressions = value.Arguments.Select(x => x.Expression).ToList();
            var lines = new List<string>
            {
                "if (!Validate(" + Literal(eventName) + "))",
                "{"
            };
            if (expressions.Count != 0) lines.Add("    " + value.InvokerExpression + ".ReleaseParamsArray(" + string.Join(", ", expressions) + ");");
            lines.Add("    return;");
            lines.Add("}");
            lines.Add("object[] paramsArray = " + (expressions.Count == 0 ? "new object[0]" : value.InvokerExpression + ".ValidateParamsArray(" + string.Join(", ", expressions) + ")") + ";");
            lines.Add("EventBinding.RaiseCustomEvent(" + Literal(eventName) + ", ref paramsArray);");
            AddWriteBack(lines, value, string.Empty);
            return string.Join("\n", lines);
        }

        private static string FactoryScalarName(WrapperInvocation value, string suffix)
        {
            var name = ScalarName(value);
            if (name == "Boolean" || name == "Bool") name = "Bool";
            if (suffix == "MethodGet")
            {
                if (name == "Float") name = "Single";
                if (name != "Int16" && name != "Int32" && name != "Double" && name != "Single" && name != "Bool" && name != "DateTime" && name != "String")
                    throw new InvalidOperationException("Factory does not provide Execute" + name + "MethodGet; use Invoker for " + value.ReturnType + ".");
            }
            else if (name == "Decimal")
            {
                throw new InvalidOperationException("Factory does not provide ExecuteDecimalPropertyGet; use Invoker for " + value.ReturnType + ".");
            }
            return name;
        }

        private static string ScalarConversion(WrapperInvocation value, string expression)
        {
            if (!string.IsNullOrWhiteSpace(value.ScalarConversion) && value.ScalarConversion.IndexOf("{0}", StringComparison.Ordinal) >= 0)
                return value.ScalarConversion.Replace("{0}", expression);
            var name = ScalarName(value);
            if (name == "Bool") name = "Boolean";
            if (name == "Float") name = "Single";
            return "NetRuntimeSystem.Convert.To" + name + "(" + expression + ")";
        }

        private static string ScalarName(WrapperInvocation value)
        {
            if (!string.IsNullOrWhiteSpace(value.ScalarConversion)) return value.ScalarConversion.Trim().Replace("global::System.Convert.To", string.Empty).Replace("NetRuntimeSystem.Convert.To", string.Empty).Replace("System.Convert.To", string.Empty).Replace("Convert.To", string.Empty);
            switch ((value.ReturnType ?? string.Empty).Trim())
            {
                case "bool": case "Boolean": case "System.Boolean": case "global::System.Boolean": return "Boolean";
                case "byte": case "Byte": case "System.Byte": case "global::System.Byte": return "Byte";
                case "short": case "Int16": case "System.Int16": case "global::System.Int16": return "Int16";
                case "int": case "Int32": case "System.Int32": case "global::System.Int32": return "Int32";
                case "long": case "Int64": case "System.Int64": case "global::System.Int64": return "Int64";
                case "float": return "Float";
                case "Single": case "System.Single": case "global::System.Single": return "Single";
                case "double": case "Double": case "System.Double": case "global::System.Double": return "Double";
                case "decimal": case "Decimal": case "System.Decimal": case "global::System.Decimal": return "Decimal";
                case "string": case "String": case "System.String": case "global::System.String": return "String";
                case "DateTime": case "System.DateTime": case "global::System.DateTime": return "DateTime";
                default: throw new InvalidOperationException("ScalarConversion is required for return type " + value.ReturnType + ".");
            }
        }

        private static string ArrayExpression(IEnumerable<string> values, WrapperObjectArrayStyle style)
        {
            return (style == WrapperObjectArrayStyle.Compact ? "new object[]{" : "new object[] {") + " " + string.Join(", ", values) + " }";
        }

        private static string JoinCall(string prefix, IEnumerable<string> tail)
        {
            var values = new[] { prefix }.Concat(tail ?? Enumerable.Empty<string>()).Where(x => !string.IsNullOrWhiteSpace(x));
            return string.Join(", ", values);
        }
        private static string Literal(string value) => "\"" + (value ?? string.Empty).Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
    }
}
