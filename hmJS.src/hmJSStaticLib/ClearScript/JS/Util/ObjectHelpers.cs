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
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;

namespace Microsoft.ClearScript.Util
{
    internal static class ObjectHelpers
    {
        // GUID_ManagedName (um\cor.h)
        private static readonly Guid managedNameGuid = new Guid("{0f21f359-ab84-41e8-9a78-36d110e6d2f9}");

        public static Type GetTypeOrTypeInfo(this object value)
        {
            var type = value.GetType();
            IDispatch dispatch = null;

            Type typeInfo = null;
            TYPEKIND typeInfoKind = 0;
            TYPEFLAGS typeInfoFlags = 0;

            if (type.IsUnknownCOMObject())
            {
                // This appears to be a generic COM object with no specific type information.
                // Attempt to acquire COM type information via IDispatch or IProvideClassInfo.

                dispatch = value as IDispatch;
                if (dispatch != null)
                {
                    uint count;
                    if (RawCOMHelpers.HResult.Succeeded(dispatch.GetTypeInfoCount(out count)) && (count > 0))
                    {
                        ITypeInfo tempTypeInfo;
                        if (RawCOMHelpers.HResult.Succeeded(dispatch.GetTypeInfo(0, 0, out tempTypeInfo)))
                        {
                            typeInfo = GetTypeForTypeInfo(tempTypeInfo);
                            typeInfoKind = GetTypeInfoKind(tempTypeInfo);
                            typeInfoFlags = GetTypeInfoFlags(tempTypeInfo);
                        }
                    }
                }

                if (typeInfo == null)
                {
                    var provideClassInfo = value as IProvideClassInfo;
                    if (provideClassInfo != null)
                    {
                        ITypeInfo tempTypeInfo;
                        if (RawCOMHelpers.HResult.Succeeded(provideClassInfo.GetClassInfo(out tempTypeInfo)))
                        {
                            typeInfo = GetTypeForTypeInfo(tempTypeInfo);
                            typeInfoKind = GetTypeInfoKind(tempTypeInfo);
                            typeInfoFlags = GetTypeInfoFlags(tempTypeInfo);
                        }
                    }
                }
            }

            if (typeInfo != null)
            {
                // If the COM type is a dispatch-only interface, use it. Such interfaces typically
                // aren't exposed via QueryInterface(), so there's no way to validate them anyway.

                if ((dispatch != null) && (typeInfoKind == TYPEKIND.TKIND_DISPATCH) && typeInfoFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FDISPATCHABLE) && !typeInfoFlags.HasFlag(TYPEFLAGS.TYPEFLAG_FDUAL))
                {
                    return typeInfo;
                }

                // COM type information acquired in this manner may not actually be valid for the
                // original object. In some cases the original object implements a base interface.

                if (typeInfo.IsInstanceOfType(value))
                {
                    return typeInfo;
                }

                foreach (var interfaceType in typeInfo.GetInterfaces())
                {
                    if (interfaceType.IsInstanceOfType(value))
                    {
                        return interfaceType;
                    }
                }
            }

            return type;
        }

        public static string GetFriendlyName(this object value)
        {
            return value.GetFriendlyName(null);
        }

        public static string GetFriendlyName(this object value, Type type)
        {
            if (type == null)
            {
                if (value == null)
                {
                    return "[null]";
                }

                type = value.GetType();
            }

            if (type.IsArray && (value != null))
            {
                var array = (Array)value;
                var dimensions = Enumerable.Range(0, type.GetArrayRank());
                var lengths = String.Join(",", dimensions.Select(array.GetLength));
                return MiscHelpers.FormatInvariant("{0}[{1}]", type.GetElementType().GetFriendlyName(), lengths);
            }

            return type.GetFriendlyName();
        }

        public static T DynamicCast<T>(this object value)
        {
            // ReSharper disable RedundantCast

            // the cast to dynamic is not redundant; removing it breaks tests
            return (T)(dynamic)value;

            // ReSharper restore RedundantCast
        }

        public static object ToDynamicResult(this object result, ScriptEngine engine)
        {
            if (result is Nonexistent)
            {
                return Undefined.Value;
            }

            if ((result is HostTarget) || (result is IPropertyBag))
            {
                // Returning an instance of HostTarget (an internal type) isn't likely to be
                // useful. Wrapping it in a dynamic object makes sense in this context. Wrapping
                // a property bag allows it to participate in dynamic invocation chaining, which
                // may be useful when dealing with things like host type collections. HostItem
                // supports dynamic conversion, so the client can unwrap the object if necessary.

                return HostItem.Wrap(engine, result);
            }

            return result;
        }

        private static Type GetTypeForTypeInfo(ITypeInfo typeInfo)
        {
            // ReSharper disable EmptyGeneralCatchClause

            try
            {
                ITypeLib typeLib;
                int index;
                typeInfo.GetContainingTypeLib(out typeLib, out index);

                var assembly = LoadPrimaryInteropAssembly(typeLib);
                if (assembly != null)
                {
                    var name = GetManagedTypeInfoName(typeInfo, typeLib);
                    var guid = GetTypeInfoGuid(typeInfo);

                    var type = assembly.GetType(name, false /*throwOnError*/, true /*ignoreCase*/);
                    if ((type != null) && (type.GUID == guid))
                    {
                        return type;
                    }

                    var types = assembly.GetTypes();
                    if ((index >= 0) && (index < types.Length))
                    {
                        type = types[index];
                        if ((type.GUID == guid) && (type.FullName == name))
                        {
                            return type;
                        }
                    }

                    type = types.FirstOrDefault(testType => (testType.GUID == guid) && (testType.FullName.Equals(name, StringComparison.OrdinalIgnoreCase)));
                    if (type != null)
                    {
                        return type;
                    }
                }

                var pTypeInfo = Marshal.GetComInterfaceForObject(typeInfo, typeof(ITypeInfo));
                try
                {
                    return Marshal.GetTypeForITypeInfo(pTypeInfo);
                }
                finally
                {
                    Marshal.Release(pTypeInfo);
                }
            }
            catch (Exception)
            {
            }

            return null;

            // ReSharper restore EmptyGeneralCatchClause
        }

        private static Assembly LoadPrimaryInteropAssembly(ITypeLib typeLib)
        {
            // ReSharper disable EmptyGeneralCatchClause

            try
            {
                IntPtr pAttr;
                typeLib.GetLibAttr(out pAttr);
                try
                {
                    var attr = (TYPELIBATTR)Marshal.PtrToStructure(pAttr, typeof(TYPELIBATTR));

                    string name;
                    string codeBase;
                    if (new TypeLibConverter().GetPrimaryInteropAssembly(attr.guid, attr.wMajorVerNum, attr.wMinorVerNum, attr.lcid, out name, out codeBase))
                    {
                        return Assembly.Load(new AssemblyName(name) { CodeBase = codeBase });
                    }
                }
                finally
                {
                    typeLib.ReleaseTLibAttr(pAttr);
                }
            }
            catch (Exception)
            {
            }

            return null;

            // ReSharper restore EmptyGeneralCatchClause
        }

        private static string GetManagedTypeInfoName(ITypeInfo typeInfo, ITypeLib typeLib)
        {
            var typeInfo2 = typeInfo as ITypeInfo2;
            if (typeInfo2 != null)
            {
                // ReSharper disable EmptyGeneralCatchClause

                try
                {
                    var guid = managedNameGuid;
                    object data;
                    typeInfo2.GetCustData(ref guid, out data);

                    var name = data as string;
                    if (name != null)
                    {
                        return name.Trim();
                    }
                }
                catch (Exception)
                {
                }

                // ReSharper restore EmptyGeneralCatchClause
            }

            return GetManagedTypeLibName(typeLib) + "." + Marshal.GetTypeInfoName(typeInfo);
        }

        private static string GetManagedTypeLibName(ITypeLib typeLib)
        {
            var typeLib2 = typeLib as ITypeLib2;
            if (typeLib2 != null)
            {
                // ReSharper disable EmptyGeneralCatchClause

                try
                {
                    var guid = managedNameGuid;
                    object data;
                    typeLib2.GetCustData(ref guid, out data);

                    var name = data as string;
                    if (name != null)
                    {
                        name = name.Trim();
                        if (name.EndsWith(".dll", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                        {
                            return name.Substring(0, name.Length - 4);
                        }

                        return name;
                    }
                }
                catch (Exception)
                {
                }

                // ReSharper restore EmptyGeneralCatchClause
            }

            return Marshal.GetTypeLibName(typeLib);
        }

        private static Guid GetTypeInfoGuid(ITypeInfo typeInfo)
        {
            IntPtr pAttr;
            typeInfo.GetTypeAttr(out pAttr);
            try
            {
                var attr = (TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(TYPEATTR));
                return attr.guid;
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(pAttr);
            }
        }

        private static TYPEKIND GetTypeInfoKind(ITypeInfo typeInfo)
        {
            IntPtr pAttr;
            typeInfo.GetTypeAttr(out pAttr);
            try
            {
                var attr = (TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(TYPEATTR));
                return attr.typekind;
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(pAttr);
            }
        }

        private static TYPEFLAGS GetTypeInfoFlags(ITypeInfo typeInfo)
        {
            IntPtr pAttr;
            typeInfo.GetTypeAttr(out pAttr);
            try
            {
                var attr = (TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(TYPEATTR));
                return attr.wTypeFlags;
            }
            finally
            {
                typeInfo.ReleaseTypeAttr(pAttr);
            }
        }

        #region Nested type: IProvideClassInfo

        [ComImport]
        [Guid("b196b283-bab4-101a-b69c-00aa00341d07")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IProvideClassInfo
        {
            [PreserveSig]
            int GetClassInfo(
                [Out] [MarshalAs(UnmanagedType.Interface)] out ITypeInfo typeInfo
            );
        }

        #endregion
    }
}
