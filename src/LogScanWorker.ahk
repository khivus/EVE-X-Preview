#Requires AutoHotkey v2.0
#SingleInstance Off
#NoTrayIcon
; This file is also embedded as an alternate script in the standalone EXE.
; Background mode lowers disk I/O and CPU priority as well as keeping UI isolated.
if !DllCall("SetPriorityClass", "Ptr", DllCall("GetCurrentProcess", "Ptr"), "UInt", 0x100000)
    try ProcessSetPriority("BelowNormal")
OnError(WorkerError)
if A_Args.Length != 2
    ExitApp(1)
jobDirectory := A_Args[1]
parentPid := Integer(A_Args[2])
parentWatch := (*) => (ProcessExist(parentPid) ? 0 : ExitApp())
SetTimer(parentWatch, 1000)
try {
    request := StrSplit(FileRead(jobDirectory "\input.txt", "UTF-8"), "`n", "`r")
    mode := request[1], directory := request[2]
    if mode = "game"
        output := ScanGame(directory)
    else if mode = "chat"
        output := ScanChat(directory, request, jobDirectory)
    else
        throw Error("Unknown log scan mode")
    PublishResult(jobDirectory, output)
    ExitApp()
} catch as err {
    WorkerError(err)
}

WorkerError(err, *) {
    if A_Args.Length
        try FileAppend(err.Message " at " err.File ":" err.Line, A_Args[1] "\error.txt", "UTF-8")
    ExitApp(1)
}

ScanGame(directory) {
    newest := Map()
    Loop Files, directory "\*.txt", "F" {
        if !RegExMatch(A_LoopFileName, "i)^(\d{8}_\d{6})_(\d+)\.txt$", &match)
            continue
        id := match[2], session := match[1]
        if !newest.Has(id) || StrCompare(session, newest[id].session) > 0
            newest[id] := {path: A_LoopFileFullPath, session: session}
    }
    output := ""
    for id, entry in newest {
        listener := ""
        try {
            file := FileOpen(entry.path, "r", "UTF-8")
            if file {
                try {
                    header := StrSplit(file.Read(4096), "`n", "`r")
                    if header.Length >= 4 && RegExMatch(header[3], "^\s*Listener:\s*(.+?)\s*$", &name)
                        && RegExMatch(header[4], "i)^\s*Session Started:\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*$")
                        listener := name[1]
                } finally file.Close()
            }
        }
        output .= id "`t" entry.path "`t" entry.session "`t" listener "`n"
    }
    return output
}

ScanChat(directory, request, jobDirectory) {
    active := Map(), cache := Map(), newest := Map()
    for index, character in request
        if index > 2 && character != ""
            active[character] := true
    if FileExist(jobDirectory "\cache.txt") {
        for row in StrSplit(FileRead(jobDirectory "\cache.txt", "UTF-8"), "`n", "`r") {
            fields := StrSplit(row, "`t")
            if fields.Length = 4
                cache[fields[1]] := {stamp: fields[2], listener: fields[3], session: fields[4]}
        }
    }
    nextCache := "", candidates := ""
    Loop Files, directory "\Local_*.txt", "F" {
        if !RegExMatch(A_LoopFileName, "i)^Local_\d{8}_\d{6}_\d+\.txt$")
            continue
        path := A_LoopFileFullPath
        stamp := A_LoopFileTimeCreated "|" A_LoopFileTimeModified "|" A_LoopFileSize
        candidates .= A_LoopFileName "`t" path "`t" stamp "`n"
    }
    ; Current sessions are found first, but all headers still determine final ordering.
    for row in StrSplit(Sort(candidates, "R"), "`n", "`r") {
        fields := StrSplit(row, "`t")
        if fields.Length != 3
            continue
        path := fields[2], stamp := fields[3]
        if cache.Has(path) && cache[path].stamp = stamp
            header := cache[path]
        else {
            try header := ReadChatHeader(path)
            catch
                continue
            if header.listener = "" || header.session = ""
                continue ; Incomplete/locked files are retried on the next discovery.
        }
        nextCache .= path "`t" stamp "`t" header.listener "`t" header.session "`n"
        if active.Has(header.listener) && (!newest.Has(header.listener) || StrCompare(header.session, newest[header.listener].session) > 0) {
            newest[header.listener] := {path: path, session: header.session}
            PublishResult(jobDirectory, ChatResult(newest))
        }
    }
    FileAppend(nextCache, jobDirectory "\cache.tmp", "UTF-8")
    FileMove(jobDirectory "\cache.tmp", jobDirectory "\cache.txt", 1)
    return ChatResult(newest)
}

ChatResult(newest) {
    output := ""
    for character, entry in newest {
        try seed := ReadLatestSystem(entry.path)
        catch
            continue
        output .= character "`t" entry.path "`t" entry.session "`t" seed.system "`t" seed.cursor "`t" seed.identity "`n"
    }
    return output
}

ReadChatHeader(path) {
    header := {listener: "", session: ""}
    file := FileOpen(path, "r", "UTF-8")
    if !file
        return header
    try {
        ; Bound header I/O even if an unrelated/corrupt log has a giant first line.
        text := file.Read(8192)
        for line in StrSplit(text, "`n", "`r") {
            if RegExMatch(line, "^\s*Listener:\s*(.+?)\s*$", &match)
                header.listener := match[1]
            else if RegExMatch(line, "i)^\s*Session started:\s*(\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})", &match)
                header.session := RegExReplace(match[1], "\D")
            if header.listener != "" && header.session != ""
                break
        }
    } finally file.Close()
    return header
}

ReadLatestSystem(path) {
    file := FileOpen(path, "r", "UTF-8")
    if !file
        throw Error("Chat log could not be opened")
    try {
        info := Buffer(52, 0)
        identity := ""
        if DllCall("GetFileInformationByHandle", "Ptr", file.Handle, "Ptr", info)
            identity := NumGet(info, 28, "UInt") ":" NumGet(info, 44, "UInt") ":" NumGet(info, 48, "UInt")
        cursor := file.Length, ending := cursor, prefix := "", system := ""
        utf16 := InStr(file.Encoding, "UTF-16")
        while ending > 0 {
            start := Max(0, ending - 65536)
            if utf16
                start -= Mod(start, 2)
            file.Seek(start)
            text := file.Read(utf16 ? 32768 : 65536) prefix
            ; Pick the last complete system announcement, never an unfinished line.
            lines := StrSplit(text, "`n", "`r")
            if SubStr(text, -1) != "`n"
                lines.Pop()
            for index, line in lines {
                if (index = 1 && start > 0) || StrLen(line) > 8192
                    continue
                if RegExMatch(line, "^\x{FEFF}?\s*\[\s*\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2}\s*\] EVE System > Channel changed to Local :\s*(.+?)\s*$", &match)
                    system := match[1]
            }
            if system != ""
                break
            prefix := SubStr(text, 1, 2048)
            ending := start
        }
        return {system: system, cursor: cursor, identity: identity}
    } finally file.Close()
}

PublishResult(jobDirectory, output) {
    static serial := 0
    for name in ["output.tmp", "progress.tmp"]
        if FileExist(jobDirectory "\" name)
            FileDelete(jobDirectory "\" name)
    FileAppend(output, jobDirectory "\output.tmp", "UTF-8")
    FileMove(jobDirectory "\output.tmp", jobDirectory "\output.txt", 1)
    FileAppend(String(++serial), jobDirectory "\progress.tmp", "UTF-8")
    FileMove(jobDirectory "\progress.tmp", jobDirectory "\progress.txt", 1)
}
