﻿/*
 * Copyright (c) 2016-2017 Akitsugu Komiyama
 * under the Apache License Version 2.0
 */

using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Collections;




// ★秀丸クラス
public sealed partial class hmJSDynamicLib
{
    public sealed partial class Hidemaru
    {
        public sealed class Macro
        {
            static Macro()
            {
                SetUnManagedDll();

            }

            // マクロでの問い合わせ結果系
            public interface IResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
            }

            // 問い合わせ結果系の実態。外から見えないように
            private class TResult : IResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }

                public TResult(int Result, String Message, Exception Error)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                }
            }

            // マクロ文字列の実行。複数行を一気に実行可能
            internal static int _Eval(String cmd)
            {
                if (version < 866)
                {
                    OutputDebugStream(ErrorMsg.MethodNeed866);
                    return 0;
                }

                int ret = 0;
                try
                {
                    ret = pEvalMacro(cmd);
                }
                catch (Exception e)
                {
                    OutputDebugStream(e.Message);
                }
                return ret;
            }

            // マクロ文字列の実行。複数行を一気に実行可能。
            // 文字列なら、そのまま、ぞれ以外なら、「engine.Script.R」の関数でヒアドキュメント化する。
            // function R(text){ で検索
            public static IResult Eval(Object here_document)
            {
                if (version < 866)
                {
                    OutputDebugStream(ErrorMsg.MethodNeed866);
                    TResult result = new TResult(0, "", new InvalidOperationException("HidemaruNeedVersionException"));
                    return result;
                }

                // 文字列で書いているようであれば、普通のEval代わりに使っている
                if (here_document.GetType().Name == "String")
                {
                    int _ret = _Eval((String)here_document);

                    if (_ret == 0)
                    {
                        TResult result = new TResult(_ret, "", new InvalidOperationException("HidemaruMacroEvalException"));
                        return result;
                    }
                    else
                    {
                        TResult result = new TResult(_ret, "", null);
                        return result;
                    }
                }
                else
                {
                    // OutputDebugStream(here_document.GetType().Name);
                    String cmd = engine.Script.R(here_document);

                    int ret = 0;
                    try
                    {
                        ret = pEvalMacro(cmd);
                    }
                    catch (Exception e)
                    {
                        OutputDebugStream(e.Message);
                    }

                    if (ret == 0)
                    {
                        TResult result = new TResult(ret, "", new InvalidOperationException("HidemaruMacroEvalException"));
                        return result;
                    }
                    else
                    {
                        TResult result = new TResult(ret, "", null);
                        return result;
                    }
                }

            }

            [DllImport("user32.dll", SetLastError = true)]
            static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, IntPtr szTitle);

            public static bool IsExecuting
            {
                get
                {
                    const int WM_USER = 0x400;
                    const int WM_ISMACROEXECUTING = WM_USER + 167;

                    // 875.02から存在するが、安全を見て875正式版以降とする
                    if (version >= 875.99)
                    {
                        IntPtr hWndHidemaru = WindowHandle;
                        if (hWndHidemaru != IntPtr.Zero)
                        {
                            bool cwch = SendMessage(hWndHidemaru, WM_ISMACROEXECUTING, IntPtr.Zero, IntPtr.Zero);
                            return cwch;
                        }
                    }
                    // 古い状態でも取れる。866以上なら余裕
                    else
                    {
                        IntPtr hWndHidemaru = WindowHandle;
                        if (hWndHidemaru != IntPtr.Zero)
                        {
                            IntPtr hHm32Client = FindWindowEx(hWndHidemaru, IntPtr.Zero, "HM32CLIENT", IntPtr.Zero);
                            bool cwch = SendMessage(hHm32Client, WM_ISMACROEXECUTING, IntPtr.Zero, IntPtr.Zero);
                            return cwch;
                        }

                    }

                    return false;
                }
            }

            public static IResult ExecEval(String cmd)
            {
                TResult result = new TResult(0, "", null);
                if (IsExecuting)
                {
                    result.Result = -1;
                    result.Message = "";
                    result.Error = new InvalidOperationException("HidemaruMacroIsExecutingException");
                    return result;
                }

                if (version < 875.02)
                {
                    OutputDebugStream(ErrorMsg.MethodNeed875);
                    result.Result = 0;
                    result.Message = "";
                    result.Error = new InvalidOperationException("HidemaruNeedVersionException");
                    return result;
                }

                IntPtr hWndHidemaru = WindowHandle;
                if (hWndHidemaru == IntPtr.Zero)
                {
                    result.Result = 0;
                    result.Message = "";
                    result.Error = new NullReferenceException("HidemaruWindowHandleException");
                    return result;
                }

                const int WM_USER = 0x400;
                const int WM_REMOTE_EXECMACRO_MEMORY = WM_USER + 272;

                StringBuilder sbExpression = new StringBuilder(cmd);
                StringBuilder sbRet = new StringBuilder("\x0f0f", 0x0f0f + 1); // 最初の値は帰り値のバッファー
                bool cwch = SendMessage(hWndHidemaru, WM_REMOTE_EXECMACRO_MEMORY, sbRet, sbExpression);
                if (cwch)
                {
                    result.Result = 1;
                    result.Message = sbRet.ToString();
                    result.Error = null;
                }
                else
                {
                    result.Result = 0;
                    result.Message = sbRet.ToString();
                    result.Error = new InvalidOperationException("HidemaruMacroEvalException");
                }
                return result;
            }

            private static List<KeyValuePair<string, object>> SetMacroVarAndMakeMacroKeyArray(object[] args, int base_random)
            {
                var arg_list = new List<KeyValuePair<String, Object>>();
                int cur_random = new Random().Next(Int16.MaxValue) + 1;
                foreach (var ref_value in args)
                {
                    var value = ref_value;
                    bool success = false;
                    cur_random++;
                    object normalized_arg = null;
                    // Boolean型であれば、True:1 Flase:0にマッピングする
                    if (value is bool)
                    {
                        success = true;
                        if ((bool)value == true)
                        {
                            normalized_arg = 1;
                        }
                        else
                        {
                            normalized_arg = 0;
                        }
                    }

                    if (value is string || value is StringBuilder)
                    {
                        success = true;
                        normalized_arg = value.ToString();
                    }

                    if (!success)
                    {
                        // v8の場合に、V8Arrayならば、全体が文字列もしくは、数値かにあわせて、List<String> or List<long>にすることで、hm.NETなど共通のList<***>処理へと糾合する
                        if (value.GetType().Name == "WindowsScriptItem")
                        {
                            try
                            {
                                dynamic test_dynamic_value = (dynamic)value;
                                int len = (int)(test_dynamic_value.length);

                                List<long> long_list = new List<long>();
                                List<string> string_list = new List<string>();

                                int list_count = 0;
                                for (int i = 0; i < len; i++)
                                {
                                    Object dv = test_dynamic_value[i];
                                    list_count++;
                                    if (dv is Int32 || dv is Int64 || dv is IntPtr)
                                    {
                                        long_list.Add((long)(dv));
                                    }
                                    else
                                    {
                                        string_list.Add(dv.ToString());
                                    }
                                }

                                if (list_count == long_list.Count)
                                {
                                    value = long_list;
                                }
                                else if (list_count == string_list.Count)
                                {
                                    value = string_list;
                                }
                                else
                                {
                                    System.Diagnostics.Trace.WriteLine("渡された Arrayに数値型と文字列型などの異なる型が入り混じっています");
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }

                    // 配列の場合を追加
                    if (!success)
                    {
                        if (value.GetType() == new List<int>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<int>((List<int>)value);
                        }
                        if (value.GetType() == new List<long>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<long>((List<long>)value);
                        }
                        if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<IntPtr>((List<IntPtr>)value);
                        }
                    }

                    if (!success)
                    {
                        if (value.GetType() == new List<string>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<String>((List<String>)value);
                        }
                    }
                    // 以上配列の場合を追加

                    if (!success)
                    {
                        // 32bit
                        if (IntPtr.Size == 4)
                        {
                            // まずは整数でトライ
                            Int32 itmp = 0;
                            success = Int32.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    normalized_arg = (Int32)(dtmp);
                                }

                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }

                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    normalized_arg = (Int64)(dtmp);
                                }
                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }
                    }


                    // 成功しなかった
                    if (!success)
                    {
                        normalized_arg = value.ToString();
                    }

                    if (normalized_arg is Int32 || normalized_arg is Int64)
                    {
                        string key = "#AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        hmJSDynamicLib.Hidemaru.Macro.Var(key, normalized_arg);
                    }
                    else if (normalized_arg is string)
                    {
                        string key = "$AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        hmJSDynamicLib.Hidemaru.Macro.Var(key, normalized_arg);
                    }
                    else if (value.GetType() == new List<int>().GetType() || value.GetType() == new List<long>().GetType() || value.GetType() == new List<IntPtr>().GetType())
                    {
                        string key = "$AsIntArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        if (value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(key + "[" + iix + "]", int_list[iix]);
                            }
                        }
                        else if (value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(key + "[" + iix + "]", long_list[iix]);
                            }
                        }
                        else if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(key + "[" + iix + "]", ptr_list[iix]);
                            }
                        }
                    }
                    else if (value.GetType() == new List<string>().GetType())
                    {
                        string key = "$AsStrArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        List<String> str_list = (List<String>)value;
                        for (int iix = 0; iix < str_list.Count; iix++)
                        {
                            hmJSDynamicLib.Hidemaru.Macro.Var(key + "[" + iix + "]", str_list[iix]);
                        }
                    }
                }
                return arg_list;
            }



            private static int statement_base_random = 0;

            // マクロでの問い合わせ結果系
            public interface IStatementResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
                dynamic Args { get; }
            }


            private class TStatementResult : IStatementResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public dynamic Args { get; set; }

                public TStatementResult(int Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<Object>(Args); // コピー渡し
                }
            }

            private struct ExecStateResult
            {
                public int Result;
                public string Message;
                public Exception Error;
                public List<Object> Args;
            }

            public static IStatementResult Statement(string funcname, params object[] args)
            {
                if (statement_base_random == 0)
                {
                    statement_base_random = new System.Random().Next(Int16.MaxValue) + 1;

                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, statement_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                tmpVar = null;
                int dll = iDllBindHandle;

                if (dll == 0)
                {
                    throw new NullReferenceException(ErrorMsg.NoDllBindHandle866);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「文」で実行する形
                string expression = $"{funcname} {args_string};\n";

                String invocate = ModifyFuncCallByDllType("{0}");
                String cmd = "" +
                expression +
                "##_tmp_dll_id_ret = dllfuncw( " + invocate + " \"SetTmpVar\", result);\n" +
                "##_tmp_dll_id_ret = 0;\n";

                // 実行する
                IResult ret = hmJSDynamicLib.Hidemaru.Macro.Eval(cmd);
                ExecStateResult result = new ExecStateResult();
                result.Result = ret.Result;
                result.Error = ret.Error;
                result.Message = ret.Message;
                result.Args = new List<object>();

                int macro_result = 0;
                if (IntPtr.Size == 4)
                {
                    macro_result = (Int32)tmpVar + 0; // 確実に複製を
                }
                else
                {
                    Int64 macro_result64 = (Int64)tmpVar + 0; // 確実に複製を
                    Int32 macro_result32 = (Int32)HmClamp<Int64>(macro_result64, Int32.MinValue, Int32.MaxValue);
                    macro_result = (Int32)macro_result32;
                }

                if (result.Error == null)
                {
                    result.Result = macro_result;
                }
                tmpVar = null; // クリア

                // 成否も含めて結果を入れる。
                // new TResult(ret.Result, ret.Message, ret.Error);

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(hmJSDynamicLib.Hidemaru.Macro.Var(l.Key));
                        hmJSDynamicLib.Hidemaru.Macro.Var(l.Key, 0);
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(hmJSDynamicLib.Hidemaru.Macro.Var(l.Key));
                        hmJSDynamicLib.Hidemaru.Macro.Var(l.Key, "");
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", "");
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }

                return new TStatementResult(result.Result, result.Message, result.Error, result.Args); ;
            }

            private static int funciton_base_random = 0;


            // マクロでの問い合わせ結果系
            public interface IFunctionResult
            {
                object Result { get; }
                String Message { get; }
                Exception Error { get; }
                dynamic Args { get; }
            }

            private class TFunctionResult : IFunctionResult
            {
                public object Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public dynamic Args { get; set; }

                public TFunctionResult(object Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<Object>(Args); // コピー渡し
                }
            }


            public struct ExecFuncResult
            {
                public object Result;
                public string Message;
                public Exception Error;
                public List<Object> Args;
            }

            public static IFunctionResult Function(string funcname, params object[] args)
            {
                if (funciton_base_random == 0)
                {
                    funciton_base_random = new System.Random().Next(Int16.MaxValue) + 1;
                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, statement_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「関数」で実行する形
                string expression = $"{funcname}({args_string})";

                //----------------------------------------------------------------
                ExecFuncResult result = new ExecFuncResult();
                result.Args = new List<object>();

                // 866より少ないのでこのリターンの正常性は考慮しなくても良い
                if (version < 866)
                {
                    OutputDebugStream(ErrorMsg.MethodNeed866);
                    result.Result = null;
                    result.Message = "";
                    result.Error = new InvalidOperationException("HidemaruNeedVersionException");
                    return new TFunctionResult(result.Result, result.Message, result.Error, result.Args); ;
                }

                tmpVar = null;
                int dll = iDllBindHandle;

                if (dll == 0)
                {
                    throw new NullReferenceException(ErrorMsg.NoDllBindHandle866);
                }

                String invocate = ModifyFuncCallByDllType("{0}");
                String cmd = "" +
                    "##_tmp_dll_id_ret = dllfuncw( " + invocate + " \"SetTmpVar\", " + expression + ");\n" +
                    "##_tmp_dll_id_ret = 0;\n";

                IResult eval_ret = Eval(cmd);

                bool is_exception = false;
                if (tmpVar == null)
                {
                    is_exception = true;
                    result.Result = null;
                    result.Message = "";
                    result.Error = new InvalidOperationException("HidemaruMacroEvalException");
                }

                if (!is_exception)
                {
                    Object ret = tmpVar;
                    tmpVar = null; // クリア

                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            result.Result = (Int32)ret + 0; // 確実に複製を
                            result.Message = eval_ret.Message;
                            result.Error = eval_ret.Error;
                        }
                        else
                        {
                            result.Result = (Int64)ret + 0; // 確実に複製を
                            result.Message = eval_ret.Message;
                            result.Error = eval_ret.Error;
                        }
                    }
                    else
                    {
                        result.Result = (String)ret + ""; // 確実に複製を
                        result.Message = eval_ret.Message;
                        result.Error = eval_ret.Error;
                    }

                }

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(hmJSDynamicLib.Hidemaru.Macro.Var(l.Key));
                        hmJSDynamicLib.Hidemaru.Macro.Var(l.Key, 0);
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(hmJSDynamicLib.Hidemaru.Macro.Var(l.Key));
                        hmJSDynamicLib.Hidemaru.Macro.Var(l.Key, "");
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", 0);
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            hmJSDynamicLib.Hidemaru.Macro.Var(l.Key + "[" + iix + "]", "");
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }


                return new TFunctionResult(result.Result, result.Message, result.Error, result.Args); ;
            }



            // マクロ文字列の実行。複数行を一気に実行可能
            public static Object Var(String var_name, Object value = null)
            {
                // 読み取りであれば…
                if (value == null)
                {
                    if (version < 866)
                    {
                        OutputDebugStream(ErrorMsg.MethodNeed866);
                        return null;
                    }

                    tmpVar = null;
                    int dll = iDllBindHandle;

                    if (dll == 0)
                    {
                        throw new NullReferenceException(ErrorMsg.NoDllBindHandle866);
                    }

                    String invocate = ModifyFuncCallByDllType("{0}");
                    String cmd = "" +
                        "##_tmp_dll_id_ret = dllfuncw( " + invocate + " \"SetTmpVar\", " + var_name + ");\n" +
                        "##_tmp_dll_id_ret = 0;\n";

                    _Eval(cmd);

                    if (tmpVar == null)
                    {
                        return null;
                    }
                    Object ret = tmpVar;
                    tmpVar = null; // クリア

                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            return (Int32)ret + 0; // 確実に複製を
                        }
                        else
                        {
                            return (Int64)ret + 0; // 確実に複製を
                        }
                    }
                    else
                    {
                        return (String)ret + ""; // 確実に複製を
                    }

                    // 書き込みであれば…
                }
                else
                {
                    // 設定先の変数が数値型
                    if (var_name.StartsWith("#"))
                    {
                        if (version < 866)
                        {
                            OutputDebugStream(ErrorMsg.MethodNeed866);
                            return null;
                        }

                        int dll = iDllBindHandle;

                        if (dll == 0)
                        {
                            throw new NullReferenceException(ErrorMsg.NoDllBindHandle866);
                        }

                        Object result = new Object();

                        // Boolean型であれば、True:1 Flase:0にマッピングする
                        if (value.GetType().Name == "Boolean")
                        {
                            if ((Boolean)value == true)
                            {
                                value = 1;
                            }
                            else
                            {
                                value = 0;
                            }
                        }

                        // 32bit
                        if (IntPtr.Size == 4)
                        {
                            // まずは整数でトライ
                            Int32 itmp = 0;
                            bool success = Int32.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    result = (Int32)(dtmp);
                                }
                                else
                                {
                                    result = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            bool success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    result = (Int64)(dtmp);
                                }
                                else
                                {
                                    result = 0;
                                }
                            }
                        }

                        SetTmpVar(result);
                        String invocate = ModifyFuncCallByDllType("{0}");
                        String cmd = "" +
                            var_name + " = dllfuncw( " + invocate + " \"PopNumVar\" );\n";
                        _Eval(cmd);
                        SetTmpVar(null);

                        return result;
                    }

                    else // if (var_name.StartsWith("$")
                    {
                        if (version < 866)
                        {
                            OutputDebugStream(ErrorMsg.MethodNeed866);
                            return null;
                        }

                        int dll = iDllBindHandle;

                        if (dll == 0)
                        {
                            throw new NullReferenceException(ErrorMsg.NoDllBindHandle866);
                        }

                        String result = value.ToString();
                        SetTmpVar(result);
                        String invocate = ModifyFuncCallByDllType("{0}");
                        String cmd = "" +
                            var_name + " = dllfuncstrw( " + invocate + " \"PopStrVar\" );\n";
                        _Eval(cmd);
                        SetTmpVar(null);

                        return result;
                    }
                }

            }
        }
    }
}
