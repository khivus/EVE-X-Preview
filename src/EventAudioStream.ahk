#Include "EventAudioOutput.ahk"
; Windows Media Foundation decodes short chunks; XAudio2 mixes them with real
; gain above 1.0. No file copies, bundled codecs or AHK callbacks on audio threads.
class EventAudioStream {
    reader := 0
    output := 0
    voice := 0
    mediaStarted := false
    chunks := []
    queuedBytes := 0
    endOfStream := false
    static maxQueuedBytes := 2 * 1024 * 1024
    static maxSampleBytes := 4 * 1024 * 1024
    ; A borrowed handle is available only while this stream holds its owner.
    engine => IsObject(this.output) ? this.output.engine : 0

    __New(path, volume) {
        partial := actual := waveFormat := 0
        success := false
        try {
            DllCall("mfplat\MFStartup", "UInt", 0x20070, "UInt", 1, "HRESULT")
            this.mediaStarted := true
            reader := 0
            DllCall("mfreadwrite\MFCreateSourceReaderFromURL", "WStr", path, "Ptr", 0, "Ptr*", &reader, "HRESULT")
            this.reader := reader
            ComCall(4, reader, "UInt", 0xFFFFFFFE, "Int", false) ; All streams off.
            ComCall(4, reader, "UInt", 0xFFFFFFFD, "Int", true) ; First audio only.
            DllCall("mfplat\MFCreateMediaType", "Ptr*", &partial, "HRESULT")
            ComCall(24, partial, "Ptr", this.Guid("48EBA18E-F8C9-4687-BF11-0A74C9F96A8F"),
                "Ptr", this.Guid("73647561-0000-0010-8000-00AA00389B71")) ; Audio.
            ComCall(24, partial, "Ptr", this.Guid("F7E34C9A-42E8-4714-B74B-CB29D72C35E5"),
                "Ptr", this.Guid("00000001-0000-0010-8000-00AA00389B71")) ; PCM.
            ComCall(7, reader, "UInt", 0xFFFFFFFD, "Ptr", 0, "Ptr", partial)
            ComCall(6, reader, "UInt", 0xFFFFFFFD, "Ptr*", &actual)
            formatSize := 0
            DllCall("mfplat\MFCreateWaveFormatExFromMFMediaType", "Ptr", actual,
                "Ptr*", &waveFormat, "UInt*", &formatSize, "UInt", 0, "HRESULT")
            this.bufferTarget := Min(EventAudioStream.maxQueuedBytes, Max(65536, NumGet(waveFormat, 8, "UInt")))
            this.output := EventAudioOutput.Prepare()
            voice := 0
            ComCall(5, this.output.engine, "Ptr*", &voice, "Ptr", waveFormat, "UInt", 0, "Float", 2.0,
                "Ptr", 0, "Ptr", 0, "Ptr", 0)
            this.voice := voice
            this.SetVolume(volume)
            this.FillQueue()
            if !this.QueuedCount()
                throw Error("The sound file contains no playable audio.")
            ComCall(19, voice, "UInt", 0, "UInt", 0) ; Start, commit now.
            success := true
        } finally {
            if partial
                ObjRelease(partial)
            if actual
                ObjRelease(actual)
            if waveFormat
                DllCall("ole32\CoTaskMemFree", "Ptr", waveFormat)
            if !success
                this.Dispose()
        }
    }

    Guid(text) {
        value := Buffer(16, 0)
        DllCall("ole32\CLSIDFromString", "WStr", "{" text "}", "Ptr", value, "HRESULT")
        return value
    }

    QueuedCount() {
        if !this.voice
            return 0
        state := Buffer(A_PtrSize + 12, 0) ; XAudio structures use pack(1).
        ComCall(25, this.voice, "Ptr", state, "UInt", 0x100, "Int") ; GetState, no sample count.
        return NumGet(state, A_PtrSize, "UInt")
    }

    ReadChunk() {
        sample := mediaBuffer := 0
        locked := false
        try {
            flags := 0
            ComCall(9, this.reader, "UInt", 0xFFFFFFFD, "UInt", 0, "Ptr", 0,
                "UInt*", &flags, "Ptr", 0, "Ptr*", &sample)
            if flags & 0x31 ; Error or an unexpected media-format change.
                throw Error("The sound file changed format during playback.")
            if flags & 2
                this.endOfStream := true
            if !sample
                return false
            ComCall(41, sample, "Ptr*", &mediaBuffer) ; ConvertToContiguousBuffer.
            data := length := 0
            ComCall(3, mediaBuffer, "Ptr*", &data, "Ptr", 0, "UInt*", &length)
            locked := true
            if length > EventAudioStream.maxSampleBytes
                throw Error("The audio decoder returned an oversized sample.")
            if !length
                return false
            ; IXAudio2 copies this descriptor, but retains the locked audio data.
            descriptor := Buffer(28 + 2 * A_PtrSize, 0)
            NumPut("UInt", this.endOfStream ? 0x40 : 0, "UInt", length, "Ptr", data, descriptor)
            chunk := {buffer: mediaBuffer, bytes: length}
            this.chunks.Push(chunk)
            try ComCall(21, this.voice, "Ptr", descriptor, "Ptr", 0) ; SubmitSourceBuffer.
            catch {
                this.chunks.Pop()
                throw
            }
            this.queuedBytes += length
            mediaBuffer := 0 ; The queue now owns the locked buffer reference.
            return true
        } finally {
            if mediaBuffer {
                if locked
                    ComCall(4, mediaBuffer) ; Unlock.
                ObjRelease(mediaBuffer)
            }
            if sample
                ObjRelease(sample)
        }
    }

    ReleaseChunk(chunk) {
        ComCall(4, chunk.buffer) ; IMFMediaBuffer.Unlock.
        ObjRelease(chunk.buffer)
        this.queuedBytes -= chunk.bytes
    }

    FillQueue() {
        queued := this.QueuedCount()
        while this.chunks.Length > queued
            this.ReleaseChunk(this.chunks.RemoveAt(1))
        started := A_TickCount
        attempts := 0
        while !this.endOfStream && this.chunks.Length < 32 && this.queuedBytes < this.bufferTarget {
            if !this.ReadChunk() && !this.endOfStream && ++attempts >= 16
                throw Error("The audio decoder returned no samples.")
            if A_TickCount - started >= 10
                break ; Keep UI work responsive, even with compressed files.
        }
    }

    IsPlaying() => this.voice && (!this.endOfStream || this.QueuedCount() > 0)

    SetVolume(volume) => ComCall(12, this.voice, "Float", volume / 100, "UInt", 0)

    NativeVolume() {
        gain := 0.0
        ComCall(13, this.voice, "Float*", &gain, "Int")
        return gain * 100
    }

    Dispose() {
        ; DestroyVoice waits for the mixer to stop touching submitted samples.
        if this.voice {
            ComCall(18, this.voice, "Int")
            this.voice := 0
        }
        for chunk in this.chunks
            this.ReleaseChunk(chunk)
        this.chunks := []
        if this.reader {
            ObjRelease(this.reader)
            this.reader := 0
        }
        if this.mediaStarted {
            DllCall("mfplat\MFShutdown", "HRESULT")
            this.mediaStarted := false
        }
        this.output := 0
    }
}
