function require(filepath) {
    var m_file_path = "";
    var m_currentmacrodirectory = hm.Macro.Var("currentmacrodirectory");
    if (clr.System.IO.File.Exists(m_currentmacrodirectory + "\\" + path + ".js")) {
        m_file_path = m_currentmacrodirectory + "\\" + path + ".js";
    } else if (clr.System.IO.File.Exists(path + ".js")) {
        m_file_path = path + ".js";
    }

    var module_code = clr.System.IO.File.ReadAllText(m_file_path);
    var eval_code = "(function(){ var exports = {};" +
        module_code + ";" +
        "return exports; })()";
    return eval(eval_code);
}