#Requires AutoHotkey v2.0

#Include Sm30JsonConfig.ahk
#Include CodeReviewAppPaths.ahk

; Load dedicated SE01 code review flow definitions from JSON.
class CodeReviewConfig {
    static LoadFile(jsonPath) {
        if (!FileExist(jsonPath)) {
            throw Error("Review config not found: " jsonPath)
        }
        jsonText := Sm30JsonConfig._ReadFileText(jsonPath)
        reviewDef := Sm30JsonConfig.LoadText(jsonText)
        CodeReviewConfig._Validate(reviewDef, jsonPath)
        return reviewDef
    }

    static LoadDefault() {
        return CodeReviewConfig.LoadFile(CodeReviewAppPaths.DefaultReviewConfigPath())
    }

    static _Validate(reviewDef, jsonPath) {
        if (Type(reviewDef) != "Object") {
            throw Error("Review config root must be an object: " jsonPath)
        }
        required := ["id", "label", "checkpoint0", "step1", "step2", "step3", "step4", "checkpoints"]
        for fieldName in required {
            if (!Sm30JsonConfig._HasProp(reviewDef, fieldName)) {
                throw Error("Missing required field '" fieldName "' in " jsonPath)
            }
        }
    }
}
