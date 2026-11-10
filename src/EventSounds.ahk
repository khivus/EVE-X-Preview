#Include "EventAudioStream.ahk"
#Include "EventSoundMixer.ahk"

; The library stores references to user files, shared across profiles.
; Removing a reference never deletes the file or changes event assignments.
class EventSounds {
    static maxVolume := 300
    static maxPlaybackVolume := 2700 ; Three 300% gains multiplied together.
    static maxSimultaneousSounds := 16
    static maxRepeatTimeout := 3600000
    static Library(config) {
        if !config.Has("SoundLibrary") || !(config["SoundLibrary"] is Array)
            config["SoundLibrary"] := []
        return config["SoundLibrary"]
    }

    static FullPath(path) {
        path := Trim(path, " `t`r`n")
        if path = ""
            return ""
        size := DllCall("GetFullPathNameW", "Str", path, "UInt", 0, "Ptr", 0, "Ptr", 0, "UInt")
        pathBuffer := Buffer(size * 2, 0)
        if !size || !DllCall("GetFullPathNameW", "Str", path, "UInt", size, "Ptr", pathBuffer, "Ptr", 0, "UInt")
            throw Error("Cannot resolve the sound file path.")
        return StrGet(pathBuffer)
    }

    static Validate(path) {
        if path = "" || InStr(path, '"') || InStr(path, "`r") || InStr(path, "`n")
            throw Error("Select a WAV or MP3 sound file.")
        SplitPath(path,,, &extension)
        if !RegExMatch(extension, "i)^(wav|mp3)$")
            throw Error("Select a WAV or MP3 sound file.")
        attributes := FileExist(path)
        if !attributes || InStr(attributes, "D")
            throw Error("Sound file not found. Select it again if it was moved or deleted.")
    }

    static Add(config, path) {
        path := this.FullPath(path)
        this.Validate(path)
        library := this.Library(config)
        for saved in library
            if saved is String && StrLower(saved) = StrLower(path)
                return saved
        library.Push(path)
        return path
    }

    static Remove(config, path) {
        library := this.Library(config)
        count := library.Length
        loop count {
            index := count - A_Index + 1
            if library[index] is String && StrLower(library[index]) = StrLower(path)
                library.RemoveAt(index)
        }
    }

    static AnyEnabled(events) {
        for event, config in events
            if config["enabled"] && config.Get("sound", "") != ""
                return true
        return false
    }

    static Percent(value) => IsNumber(value) ? Max(0, Min(this.maxVolume, Round(value))) : 100

    static SameFile(first, second) => StrLower(first) = StrLower(second)

    static MasterVolume(config) => this.Percent(config.Get("MasterSoundVolume", 100))

    static MaxSimultaneous(config) {
        value := config.Get("MaxSimultaneousSounds", 1)
        return IsNumber(value) ? Max(1, Min(this.maxSimultaneousSounds, Round(value))) : 1
    }

    static RepeatTimeout(config) {
        value := config.Get("SoundRepeatTimeout", 2000)
        return IsNumber(value) ? Max(0, Min(this.maxRepeatTimeout, Round(value))) : 2000
    }

    static Volumes(config) {
        if !config.Has("SoundVolumes") || !(config["SoundVolumes"] is Map)
            config["SoundVolumes"] := Map()
        return config["SoundVolumes"]
    }

    static SoundVolume(config, path) => this.Percent(this.Volumes(config).Get(StrLower(path), 100))

    static EffectiveVolume(config, path, eventConfig?) {
        eventVolume := IsSet(eventConfig) ? this.Percent(eventConfig.Get("soundVolume", 100)) : 100
        return this.MasterVolume(config) * this.SoundVolume(config, path) * eventVolume / 10000
    }

    static AssignmentPlan(events, path, checked) {
        plan := {changes: Map(), replacements: [], removals: []}
        if path = ""
            return plan
        for event, config in events {
            current := config.Get("sound", "")
            if checked.Get(event, false) {
                if this.SameFile(current, path)
                    continue
                plan.changes[event] := path
                if current != ""
                    plan.replacements.Push(event)
            } else if this.SameFile(current, path) {
                plan.changes[event] := ""
                plan.removals.Push(event)
            }
        }
        return plan
    }
}

; One asynchronous Windows audio channel. Close completed media promptly so user
; files are not left locked. No player windows, bundled sounds or fallback beeps.
class EventSoundPlayer {
    activePath := ""
    activeVolume := 100
    lastError := ""
    stream := 0

    __New() {
        this.poll := ObjBindMethod(this, "CheckFinished")
        this.onExit := ObjBindMethod(this, "Stop")
        OnExit(this.onExit)
        EventAudioOutput.WarmUp()
    }

    Play(path, restart := false, volume := 100) {
        previousCritical := A_IsCritical
        Critical("On") ; Do not replace/dispose a decoder halfway through opening it.
        this.lastError := ""
        try {
            volume := Max(0, Min(EventSounds.maxPlaybackVolume, volume))
            if !volume {
                this.Stop()
                return false
            }
            EventSounds.Validate(path)
            if !restart && EventSounds.SameFile(path, this.activePath) && this.IsPlaying()
                return false ; Several characters can report the same alert in one poll.
            this.Stop()
            this.stream := EventAudioStream(path, volume)
            this.activePath := path
            this.activeVolume := volume
            SetTimer(this.poll, 20)
            return true
        } catch as err {
            this.Stop()
            this.lastError := err.Message
            return false
        } finally {
            Critical(previousCritical)
        }
    }

    SetVolume(volume) {
        previousCritical := A_IsCritical
        Critical("On")
        try {
            volume := Max(0, Min(EventSounds.maxPlaybackVolume, volume))
            if IsObject(this.stream)
                this.stream.SetVolume(volume)
            this.activeVolume := volume
        } finally Critical(previousCritical)
    }

    IsPlaying() => IsObject(this.stream) && this.stream.IsPlaying()

    NativeVolume() => IsObject(this.stream) ? this.stream.NativeVolume() : 0

    CheckFinished(*) {
        previousCritical := A_IsCritical
        Critical("On") ; A timer must not tear down buffers while another refills them.
        try {
            this.stream.FillQueue()
            if this.IsPlaying()
                return
            this.Stop()
        } catch as err {
            this.lastError := err.Message
            this.Stop()
        } finally Critical(previousCritical)
    }

    Stop(*) {
        previousCritical := A_IsCritical
        Critical("On")
        try {
            SetTimer(this.poll, 0)
            if IsObject(this.stream) {
                this.stream.Dispose()
                this.stream := 0
            }
            this.activePath := ""
        } finally Critical(previousCritical)
    }

    Dispose() {
        this.Stop()
        OnExit(this.onExit, 0)
    }
}
