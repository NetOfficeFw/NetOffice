using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Reflection.Emit;
using System.Linq.Expressions;

namespace Quack
{
    public class CodeGenUtil
    {
        public static AssemblyBuilder CreateAssemblyBuilder(string assemblyName)
        {
            return AppDomain.CurrentDomain.DefineDynamicAssembly(
                 new AssemblyName(assemblyName),
                                //AssemblyBuilderAccess.Run

#if DEBUG
                                AssemblyBuilderAccess.RunAndSave
#else
                                AssemblyBuilderAccess.Run
#endif
            );
        }

        public static ModuleBuilder CreateModuleBuilder(AssemblyBuilder assemblyBuilder, string moduleName)
        {
            //return assemblyBuilder.DefineDynamicModule(moduleName);

#if DEBUG
            return assemblyBuilder.DefineDynamicModule(moduleName, moduleName + ".dll", true);
#else
            return assemblyBuilder.DefineDynamicModule(modeleName);
#endif
        }

        public static ConstructorInfo GetConstructorInfo<T>(Expression<Func<T>> expression)
        {
            var body = expression.Body as NewExpression;
            if (body == null)
                throw new InvalidOperationException("Invalid expression form passed");

            return body.Constructor;
        }

        public static readonly OpCode[] ArgsOpcodes = {
            OpCodes.Ldarg_1,
            OpCodes.Ldarg_2,
            OpCodes.Ldarg_3
        };

        public static void EmitLoadArgument(ILGenerator il, int argumentNumber)
        {
            if (argumentNumber < ArgsOpcodes.Length)
            {
                il.Emit(ArgsOpcodes[argumentNumber]);
            }
            else
            {
                il.Emit(OpCodes.Ldarg, argumentNumber + 1);
            }
        }

        public static MethodBuilder CreateDelegateImplementation(TypeBuilder typeBuilder, FieldBuilder targetField, MethodInfo mi)
        {
            MethodBuilder methodBuilder = typeBuilder.DefineMethod(mi.Name,
                MethodAttributes.Public | MethodAttributes.Virtual,
                mi.ReturnType,
                mi.GetParameters().Select(param => param.ParameterType).ToArray());

            ILGenerator il = methodBuilder.GetILGenerator();

            #region forwarding implementation

            LocalBuilder baseReturn = null;

            if (mi.ReturnType != typeof(void))
            {
                baseReturn = il.DeclareLocal(mi.ReturnType);
            }

            // Call the target method
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, targetField);

            // Load the call parameters
            for (int i = 0; i < mi.GetParameters().Length; i++)
            {
                CodeGenUtil.EmitLoadArgument(il, i);
            }

            // Make the call
            MethodInfo callTarget = targetField.FieldType.GetMethod(mi.Name, mi.GetParameters().Select(pi => pi.ParameterType).ToArray());
            il.Emit(OpCodes.Callvirt, callTarget);

            if (mi.ReturnType != typeof(void))
            {
                il.Emit(OpCodes.Stloc_0);
                il.Emit(OpCodes.Ldloc_0);
            }

            il.Emit(OpCodes.Ret);

            #endregion

            return methodBuilder;
        }

    }
}
