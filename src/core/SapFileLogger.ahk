#Requires AutoHotkey v2.0

#Include SapHookPolicy.ahk

class SapFileLogger {
    __New(logPath := "") {
        if (logPath = "") {
            logPath := A_ScriptDir "\..\logs\sm30_" FormatTime(, "yyyyMMdd_HHmmss") ".log"
        }
        this.logPath := logPath
        SplitPath(this.logPath, , &dir)
        if (dir != "") {
            DirCreate(dir)
        }
        this.Write("INFO", "Log started: " this.logPath)
    }

    Info(message) {
        this.Write("INFO", message)
    }

    Warn(message) {
        this.Write("WARN", message)
    }

    Error(message) {
        this.Write("ERROR", message)
    }

    Write(level, message) {
        timestamp := FormatTime(, "yyyy-MM-dd HH:mm:ss")
        line := timestamp " [" level "] " message
        try {
            FileAppend(line "`n", this.logPath, "UTF-8")
        } catch {
            OutputDebug(line)
        }
    }
}

class LoggingSapHookPolicy extends SapHookPolicy {
    __New(logger := "") {
        this.logger := logger
    }

    On_Call(op, typeName, member, path, args) {
        if (!IsObject(this.logger)) {
            return
        }
        this.logger.Info("COM " op " " typeName "." member " @ " path " args=" SapLogFormat.Args(args))
    }

    After_Call(op, typeName, member, path, result) {
        if (!IsObject(this.logger)) {
            return
        }
        this.logger.Info("COM ok " typeName "." member " @ " path " => " SapLogFormat.Result(result))
    }

    On_Error(op, typeName, member, path, args) {
        if (IsObject(this.logger)) {
            this.logger.Error("COM fail " op " " typeName "." member " @ " path
                . " args=" SapLogFormat.Args(args)
                . " LastError=" A_LastError)
        }
        super.On_Error(op, typeName, member, path, args)
    }
}

class SapLogFormat {
    static Args(args) {
        if (!IsObject(args)) {
            return "[]"
        }
        parts := []
        for arg in args {
            parts.Push(this.Value(arg))
        }
        return "[" this.Join(parts, ", ") "]"
    }

    static Result(result) {
        if (IsObject(result)) {
            if (result.HasProp("_typeName")) {
                return "COM:" result._typeName
            }
            return "<object>"
        }
        return this.Value(result)
    }

    static Value(value) {
        if (IsObject(value)) {
            return "<object>"
        }
        text := String(value)
        if (StrLen(text) > 200) {
            return SubStr(text, 1, 200) "..."
        }
        return text
    }

    static Join(parts, delimiter := ", ") {
        output := ""
        for part in parts {
            if (output != "") {
                output .= delimiter
            }
            output .= part
        }
        return output
    }
}
