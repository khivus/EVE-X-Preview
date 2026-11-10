; One process-wide output device; individual streams own their source voices.
; Holding an output reference until after DestroyVoice preserves native order.
class EventAudioOutput {
    static shared := 0
    static preparing := false
    engine := 0
    mastering := 0
    module := 0

    static Prepare() {
        previousCritical := A_IsCritical
        Critical("On")
        try {
            if IsObject(this.shared)
                return this.shared
            if this.preparing
                throw Error("Audio output is still initializing. Try playback again.")
            this.preparing := true
            try {
                ; Publish only after both native resources are ready. A failed
                ; preparation must leave the next playback free to retry.
                this.shared := EventAudioOutput()
                return this.shared
            } finally this.preparing := false
        } finally Critical(previousCritical)
    }

    static WarmUp() {
        try {
            this.Prepare()
            return true
        } catch {
            ; A missing audio device must not block application/settings startup.
            ; Play reports the native error if output is still unavailable.
            return false
        }
    }

    __New() {
        success := false
        try {
            this.module := DllCall("LoadLibraryW", "Str", "xaudio2_9.dll", "Ptr")
            if !this.module
                throw Error("Windows XAudio2 is unavailable.")
            engine := 0
            ; XAUDIO2_DEFAULT_PROCESSOR (Processor1). The all-bits mask creates
            ; a worker for every processor instead of selecting the default.
            DllCall("xaudio2_9\XAudio2Create", "Ptr*", &engine, "UInt", 0, "UInt", 1, "HRESULT")
            this.engine := engine
            mastering := 0
            ; Keep the virtual default endpoint, including automatic device
            ; switching. Opening this voice is startup work, not per-alert work.
            ComCall(7, engine, "Ptr*", &mastering, "UInt", 0, "UInt", 0, "UInt", 0,
                "Ptr", 0, "Ptr", 0, "Int", 6) ; GameEffects.
            this.mastering := mastering
            success := true
        } finally {
            if !success
                this.Dispose()
        }
    }

    __Delete() => this.Dispose()

    Dispose() {
        if this.mastering {
            ComCall(18, this.mastering, "Int")
            this.mastering := 0
        }
        if this.engine {
            ObjRelease(this.engine)
            this.engine := 0
        }
        if this.module {
            DllCall("FreeLibrary", "Ptr", this.module)
            this.module := 0
        }
    }
}
