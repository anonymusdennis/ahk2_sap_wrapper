#Requires AutoHotkey v2.0

; Resolve paths for script mode and compiled exe mode.
; Compiled: config/data/logs live next to the exe (A_ScriptDir).
class Sm30AppPaths {
    static BaseDir() {
        return A_ScriptDir
    }

    static ConfigDir() {
        return Sm30AppPaths.BaseDir() "\config"
    }

    static TablesDir() {
        return Sm30AppPaths.ConfigDir() "\tables"
    }

    static DataDir() {
        return Sm30AppPaths.BaseDir() "\data"
    }

    static LogsDir() {
        return Sm30AppPaths.BaseDir() "\logs"
    }

    static IsCompiled() {
        return A_IsCompiled
    }
}
