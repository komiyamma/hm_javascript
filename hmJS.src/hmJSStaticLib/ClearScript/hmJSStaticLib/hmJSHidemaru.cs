/*
 * Copyright (c) 2016-2017 Akitsugu Komiyama
 * under the Apache License Version 2.0
 */

using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Microsoft.ClearScript;




// ★秀丸クラス
public sealed partial class hmJSDynamicLib
{
    public sealed partial class Hidemaru
    {
        public Hidemaru()
        {
            System.Diagnostics.FileVersionInfo vi = System.Diagnostics.FileVersionInfo.GetVersionInfo(strExecuteFullpath);
            _ver = 100 * vi.FileMajorPart + 10 * vi.FileMinorPart + 1 * vi.FileBuildPart + 0.01 * vi.FilePrivatePart;
            SetUnManagedDll();
        }

        public sealed class ErrorMsg
        {
            public const String MethodNeed866 = "このメソッドは秀丸エディタ v8.66 正式版以降で利用可能です。";
            public const String MethodNeed873 = "このメソッドは秀丸エディタ v8.73 正式版以降で利用可能です。";
            public const String MethodNeed875 = "このメソッドは秀丸エディタ v8.75 正式版以降で利用可能です。";
            public const String MethodNeed877 = "このメソッドは秀丸エディタ v8.77 正式版以降で利用可能です。";
            public const String MethodNeed890 = "このメソッドは秀丸エディタ v8.90 正式版以降で利用可能です。";
            public const String MethodNeed912 = "このメソッドは秀丸エディタ v9.12 正式版以降で利用可能です。";
            public const String MethodNeed915 = "このメソッドは秀丸エディタ v9.15 正式版以降で利用可能です。";
            public const String MethodNeedOutputNotFound = "HmOutputPaneの対象関数を発見できません。";
            public const String MethodNeedOutputOperation = "HmOutputPaneへの操作中にエラーが発生しました。";
            public const String MethodNeedExplorerNotFound = "HmExplorerPaneの対象関数を発見できません。";
            public const String MethodNeedExplorerOperation = "HmExplorerPaneへの操作中にエラーが発生しました。";
            public static readonly String NoDllBindHandle866 = strDllFullPath + "をloaddllした際の束縛変数の値を特定できません";

        }

        private static T HmClamp<T>(T val, T min, T max) where T : IComparable<T>
        {
            if (val.CompareTo(min) < 0) return min;
            else if (val.CompareTo(max) > 0) return max;
            else return val;
        }

        private static bool LongToInt(long number, out int intvar)
        {
            int ret_number = 0;
            while (true)
            {
                if (number > Int32.MaxValue)
                {
                    number = number - 4294967296;
                    number = number - Int32.MinValue;
                    number = number % 4294967296;
                    number = number + Int32.MinValue;
                }
                else
                {
                    break;
                }
            }
            while (true)
            {
                if (number < Int32.MinValue)
                {
                    number = number + 4294967296;
                    number = number + Int32.MinValue;
                    number = number % 4294967296;
                    number = number - Int32.MinValue;
                }
                else
                {
                    break;
                }
            }

            bool success = false;
            if (Int32.MinValue <= number && number <= Int32.MaxValue)
            {
                ret_number = (int)number;
                success = true;
            }

            intvar = ret_number;
            return success;
        }

        private static bool IsDoubleNumeric(object value)
        {
            return value is double || value is float;
        }

        private static IntPtr _hWndHidemaru = IntPtr.Zero;
        public static IntPtr WindowHandle
        {
            get
            {
                if (pGetCurrentWindowHandle != null)
                {
                    // System.Diagnostics.Trace.WriteLine("自動取得");
                    // IntPtr tmp = pGetCurrentWindowHandle();
                    // System.Diagnostics.Trace.WriteLine(tmp);
                    return pGetCurrentWindowHandle();

                }

                // System.Diagnostics.Trace.WriteLine("手動取得");
                _hWndHidemaru = HmWndHandleSearcher.GetCurWndHidemaru(_hWndHidemaru);
                // System.Diagnostics.Trace.WriteLine(_hWndHidemaru);
                if (_hWndHidemaru != IntPtr.Zero)
                {
                    return _hWndHidemaru;
                }

                return IntPtr.Zero;
            }
        }

        public static object require(string filepath)
        {
            var m_file_path = "";
            var m_currentmacrodirectory = (string)Hidemaru.Macro.Var("currentmacrodirectory");
            if (filepath.ToLower().EndsWith(".json"))
            {
                if (System.IO.File.Exists(m_currentmacrodirectory + "\\" + filepath))
                {
                    m_file_path = m_currentmacrodirectory + "\\" + filepath;
                }
                if (System.IO.File.Exists(filepath))
                {
                    if (filepath.Contains(@"\") || filepath.Contains(@"/")) { 
                        m_file_path = filepath;
                    }
                }

                if (m_file_path == "")
                {
                    var err = new System.IO.FileNotFoundException("HidemaruMacroRequireFileNotFoundException: \n" + filepath);
                    OutputDebugStream("HidemaruMacroRequireFileNotFoundException: \n" + filepath);
                    return err;
                }

                var json_data = System.IO.File.ReadAllText(m_file_path);
                return engine.Script.JSON.parse(json_data);
            }

            if (System.IO.File.Exists(m_currentmacrodirectory + "\\" + filepath + ".js"))
            {
                m_file_path = m_currentmacrodirectory + "\\" + filepath + ".js";
            }
            else if (System.IO.File.Exists(m_currentmacrodirectory + "\\" + filepath))
            {
                m_file_path = m_currentmacrodirectory + "\\" + filepath;
            }
            else if (System.IO.File.Exists(filepath + ".js"))
            {
                if (filepath.Contains(@"\") || filepath.Contains(@"/"))
                {
                    m_file_path = filepath + ".js";
                }
            }
            else if (System.IO.File.Exists(filepath))
            {
                if (filepath.Contains(@"\") || filepath.Contains(@"/"))
                {
                    m_file_path = filepath;
                }
            }

            if (m_file_path == "")
            {
                if (filepath.ToLower().EndsWith(".js"))
                {
                    var err = new System.IO.FileNotFoundException("HidemaruMacroRequireFileNotFoundException: \n" + filepath);
                    OutputDebugStream("HidemaruMacroRequireFileNotFoundException: \n" + filepath);
                    return err;
                }
                else
                {
                    var err = new System.IO.FileNotFoundException("HidemaruMacroRequireFileNotFoundException: \n" + filepath + ".js");
                    OutputDebugStream("HidemaruMacroRequireFileNotFoundException: \n" + filepath + ".js");
                    return err;
                }
            }

            var module_code = System.IO.File.ReadAllText(m_file_path);
            // exportsが空 =(Object.keys(exports).length == 0 || exports.constructor == Object) でないなら、
            // exportsを返す。それ以外は、module.exportsを返す。
            var expression = "(function(){ var module = { exports: {} }; var exports = module.exports; " +
            module_code + "; " + "\nreturn module.exports; })()";

            Object eval_obj = null;

            try
            {
                // 文字列からソース生成
                eval_obj = engine.Evaluate(expression);
            }
            catch (ScriptEngineException e)
            {
                OutputDebugStream("in " + m_file_path);
                OutputDebugStream(e.GetType().Name + ":");
                OutputDebugStream(e.ErrorDetails);

                var stack = engine.GetStackTrace();
                OutputDebugStream(stack.ToString());

                ScriptEngineException next = e.InnerException as ScriptEngineException;
                while (next != null)
                {
                    OutputDebugStream(next.ErrorDetails);
                    next = next.InnerException as ScriptEngineException;
                }
            }
            catch (ScriptInterruptedException e)
            {
                OutputDebugStream("in " + m_file_path);
                OutputDebugStream(e.GetType().Name + ":");
                OutputDebugStream(e.ErrorDetails);

                var stack = engine.GetStackTrace();
                OutputDebugStream(stack.ToString());

                ScriptInterruptedException next = e.InnerException as ScriptInterruptedException;
                while (next != null)
                {
                    OutputDebugStream(next.ErrorDetails);
                    next = next.InnerException as ScriptInterruptedException;
                }
            }
            catch (Exception e)
            {
                OutputDebugStream("in " + m_file_path);
                OutputDebugStream(e.GetType().Name + ":");
                OutputDebugStream(e.Message);
            }

            return eval_obj;
        }

        // debuginfo関数
        public static void debuginfo(params Object[] expressions)
        {
            List<String> list = new List<String>();
            foreach (var exp in expressions)
            {
                bool isClearScriptItem = false;
                try
                {
                    if (exp.GetType().Name == "WindowsScriptItem" || exp.GetType().BaseType?.Name == "WindowsScriptItem" || exp.GetType().BaseType?.BaseType?.Name == "WindowsScriptItem")
                    {
                        isClearScriptItem = true;
                    }
                }
                catch (Exception)
                {

                }
                // WindowsScriptItemエンジンのオブジェクトであれば、そのまま出しても意味が無いので…
                if (isClearScriptItem)
                {
                    dynamic dexp = exp;

                    // JSON.stringifyで変換できるはずだ
                    String strify = "";
                    try
                    {
                        strify = engine.Script.JSON.stringify(dexp);
                        list.Add(strify);
                    }
                    catch (Exception)
                    {

                    }

                    // JSON.stringfyで空っぽだったようだ。
                    if (strify == String.Empty)
                    {
                        try
                        {
                            // ECMAScriptのtoString()で変換出来るはずだ…
                            list.Add(dexp.toString());
                        }
                        catch (Exception)
                        {
                            // 変換できなかったら、とりあえず、しょうがないので.NETのToStringで。多分意味はない。
                            list.Add(exp.ToString());
                        }
                    }
                }

                // WindowsScriptItemオブジェクトでないなら、普通にToString
                else
                {
                    list.Add(exp.ToString());
                }
            }

            String joind = String.Join(" ", list);
            OutputDebugStream(joind);
        }

        // バージョン。hm.versionのため。読み取り専用
        static double _ver;
        public static double version
        {
            get { return _ver; }
        }
    }
}
