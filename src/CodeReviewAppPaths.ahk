#Requires AutoHotkey v2.0

; Paths for the code review player (script/exe lives in examples/code-review/).
class CodeReviewAppPaths {
    static BaseDir() {
        return A_ScriptDir
    }

    static ConfigDir() {
        return CodeReviewAppPaths.BaseDir() "\config"
    }

    static DataDir() {
        return CodeReviewAppPaths.BaseDir() "\data"
    }

    static LogsDir() {
        return CodeReviewAppPaths.BaseDir() "\logs"
    }

    static DefaultReviewConfigPath() {
        return CodeReviewAppPaths.ConfigDir() "\se01_review.json"
    }
}
