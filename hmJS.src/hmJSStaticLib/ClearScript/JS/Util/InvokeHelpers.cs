// 
// Copyright (c) Microsoft Corporation. All rights reserved.
// 
// Microsoft Public License (MS-PL)
// 
// This license governs use of the accompanying software. If you use the
// software, you accept this license. If you do not accept the license, do not
// use the software.
// 
// 1. Definitions
// 
//   The terms "reproduce," "reproduction," "derivative works," and
//   "distribution" have the same meaning here as under U.S. copyright law. A
//   "contribution" is the original software, or any additions or changes to
//   the software. A "contributor" is any person that distributes its
//   contribution under this license. "Licensed patents" are a contributor's
//   patent claims that read directly on its contribution.
// 
// 2. Grant of Rights
// 
//   (A) Copyright Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free copyright license
//       to reproduce its contribution, prepare derivative works of its
//       contribution, and distribute its contribution or any derivative works
//       that you create.
// 
//   (B) Patent Grant- Subject to the terms of this license, including the
//       license conditions and limitations in section 3, each contributor
//       grants you a non-exclusive, worldwide, royalty-free license under its
//       licensed patents to make, have made, use, sell, offer for sale,
//       import, and/or otherwise dispose of its contribution in the software
//       or derivative works of the contribution in the software.
// 
// 3. Conditions and Limitations
// 
//   (A) No Trademark License- This license does not grant you rights to use
//       any contributors' name, logo, or trademarks.
// 
//   (B) If you bring a patent claim against any contributor over patents that
//       you claim are infringed by the software, your patent license from such
//       contributor to the software ends automatically.
// 
//   (C) If you distribute any portion of the software, you must retain all
//       copyright, patent, trademark, and attribution notices that are present
//       in the software.
// 
//   (D) If you distribute any portion of the software in source code form, you
//       may do so only under this license by including a complete copy of this
//       license with your distribution. If you distribute any portion of the
//       software in compiled or object code form, you may only do so under a
//       license that complies with this license.
// 
//   (E) The software is licensed "as-is." You bear the risk of using it. The
//       contributors give no express warranties, guarantees or conditions. You
//       may have additional consumer rights under your local laws which this
//       license cannot change. To the extent permitted under your local laws,
//       the contributors exclude the implied warranties of merchantability,
//       fitness for a particular purpose and non-infringement.
//       

using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;

namespace Microsoft.ClearScript.Util
{
    internal static class InvokeHelpers
    {
        public static object InvokeMethod(IHostInvokeContext context, object target, MethodInfo method, object[] args)
        {
            var argList = new List<object>();
            var byRefArgInfo = new List<ByRefArgItem>();
            object tailArgsArg = null;

            var parameters = method.GetParameters();
            for (var index = 0; index < parameters.Length; index++)
            {
                var param = parameters[index];
                if (Attribute.IsDefined(param, typeof(ParamArrayAttribute)))
                {
                    if ((index != (args.Length - 1)) || !param.ParameterType.IsInstanceOfType(args[index]))
                    {
                        var tailArgType = param.ParameterType.GetElementType();
                        var tailArgs = Array.CreateInstance(tailArgType, args.Length - index);
                        for (var innerIndex = index; innerIndex < args.Length; innerIndex++)
                        {
                            var byRefArg = args[innerIndex] as IByRefArg;
                            if (byRefArg == null)
                            {
                                tailArgs.SetValue(GetCompatibleArg(param.Name, tailArgType, args[innerIndex]), innerIndex - index);
                            }
                            else
                            {
                                tailArgs.SetValue(GetCompatibleArg(param.Name, tailArgType, byRefArg.Value), innerIndex - index);
                                byRefArgInfo.Add(new ByRefArgItem(byRefArg, tailArgs, innerIndex - index));
                            }
                        }

                        argList.Add(tailArgs);
                        tailArgsArg = tailArgs;
                        break;
                    }
                }

                if (index < args.Length)
                {
                    var byRefArg = args[index] as IByRefArg;
                    if (byRefArg == null)
                    {
                        argList.Add(GetCompatibleArg(param, args[index]));
                    }
                    else
                    {
                        argList.Add(GetCompatibleArg(param, byRefArg.Value));
                        byRefArgInfo.Add(new ByRefArgItem(byRefArg, null, index));
                    }
                }
                else if (param.IsOptional)
                {
                    if (param.Attributes.HasFlag(ParameterAttributes.HasDefault))
                    {
                        try
                        {
                            argList.Add(param.DefaultValue);
                        }
                        catch (FormatException)
                        {
                            // undocumented but observed when calling HostFunctions.newVar()
                            argList.Add(null);
                        }
                    }
                    else
                    {
                        argList.Add(Missing.Value);
                    }
                }
                else
                {
                    break;
                }
            }

            var finalArgs = argList.ToArray();
            var result = method.Invoke(target, finalArgs);

            foreach (var item in byRefArgInfo)
            {
                var array = item.Array ?? finalArgs;
                item.ByRefArg.Value = array.GetValue(item.Index);
            }

            for (var index = 0; index < finalArgs.Length; index++)
            {
                if (index >= args.Length)
                {
                    break;
                }

                var finalArg = finalArgs[index];
                if (ReferenceEquals(finalArg, tailArgsArg))
                {
                    break;
                }

                args[index] = finalArg;
            }

            var type = method.ReturnType;
            if (type == typeof(void))
            {
                return VoidResult.Value;
            }

            return context.Engine.PrepareResult(result, type, method.GetScriptMemberFlags(), false);
        }

        public static object InvokeDelegate(IHostInvokeContext context, Delegate del, object[] args)
        {
            return InvokeMethod(context, del, del.GetType().GetMethod("Invoke"), args);
        }

        public static bool TryInvokeObject(IHostInvokeContext context, object target, BindingFlags invokeFlags, object[] args, object[] bindArgs, bool tryDynamic, out object result)
        {
            var hostTarget = target as HostTarget;
            if (hostTarget != null)
            {
                if (hostTarget.TryInvoke(context, invokeFlags, args, bindArgs, out result))
                {
                    return true;
                }

                if (hostTarget is HostType)
                {
                    return false;
                }

                target = hostTarget.InvokeTarget;
                tryDynamic = tryDynamic && typeof(IDynamicMetaObjectProvider).IsAssignableFrom(hostTarget.Type);
            }

            if ((target != null) && invokeFlags.HasFlag(BindingFlags.InvokeMethod))
            {
                var scriptItem = target as ScriptItem;
                if (scriptItem != null)
                {
                    target = DelegateFactory.CreateFunc<object>(scriptItem.Engine, target, args.Length);
                }

                var del = target as Delegate;
                if (del != null)
                {
                    result = InvokeDelegate(context, del, args);
                    return true;
                }

                if (tryDynamic)
                {
                    var dynamicMetaObjectProvider = target as IDynamicMetaObjectProvider;
                    if (dynamicMetaObjectProvider != null)
                    {
                        if (dynamicMetaObjectProvider.GetMetaObject(Expression.Constant(target)).TryInvoke(context, args, out result))
                        {
                            return true;
                        }
                    }
                }
            }

            result = null;
            return false;
        }

        private static object GetCompatibleArg(ParameterInfo param, object value)
        {
            return GetCompatibleArg(param.Name, param.ParameterType, value);
        }

        private static object GetCompatibleArg(string paramName, Type type, object value)
        {
            if (!type.IsAssignableFrom(ref value))
            {
                throw new ArgumentException(MiscHelpers.FormatInvariant("Invalid argument specified for parameter '{0}'", paramName));
            }

            return value;
        }

        #region Nested type: ByRefArgItem

        private class ByRefArgItem
        {
            public IByRefArg ByRefArg { get; private set; }

            public Array Array { get; private set; }

            public int Index { get; private set; }

            public ByRefArgItem(IByRefArg arg, Array array, int index)
            {
                ByRefArg = arg;
                Array = array;
                Index = index;
            }
        }

        #endregion
    }
}
