; One bounded set of per-second damage totals per character. The shared log
; reader feeds new complete lines; no extra file reads or timers are needed.
#Include "data/DamageProfiles.ahk"

class DPSMeter {
    __New(settings) {
        this.settings := settings.Clone()
        seconds := settings.Get("dpsAverageSeconds", 10)
        this.WindowSeconds := IsInteger(seconds) && seconds >= 1 && seconds <= 3600 ? seconds + 0 : 10
        this.buckets := Map()
        this.lastTimestamp := ""
        this.lastSecond := 0
        this.damageText := ""
        if settings["incomingDPSEnabled"] && settings["showIncomingDPSResistances"]
            this.profiles := DPSMeter.Profiles()
    }

    static CurrentSecond() {
        static lastUTC := "", second := 0
        utc := A_NowUTC
        if utc != lastUTC {
            second := DateDiff(utc, "19700101000000", "Seconds")
            lastUTC := utc
        }
        return second
    }

    static Profiles() {
        static profiles := 0
        if !IsObject(profiles) {
            profiles := Map()
            profiles.CaseSense := "Off"
            DamageProfiles.AddTo(profiles)
        }
        return profiles
    }

    DamageMix(line) {
        static missileMixes := Map("mjolnir", [1, 0, 0, 0], "inferno", [0, 1, 0, 0], "scourge", [0, 0, 1, 0], "nova", [0, 0, 0, 1])
        static smartbombMixes := Map("emp", [1, 0, 0, 0], "plasma", [0, 1, 0, 0], "graviton", [0, 0, 1, 0], "proton", [0, 0, 0, 1], "multispectrum", [0.25, 0.25, 0.25, 0.25])
        plain := RegExReplace(line, "<[^>]*>", "")
        if !RegExMatch(plain, "i)\(combat\)\s+[\d.]+\s+from\s+(.+?)(?:\s+-\s+(.*))?$", &source)
            return 0
        parts := StrSplit(source[2], " - ")
        weapon := parts.Length ? Trim(parts[1]) : ""
        if this.profiles.Has(weapon)
            return this.profiles[weapon].mix
        if RegExMatch(weapon, "i)\b(EMP|Plasma|Graviton|Proton|Multispectrum)\s+Smartbomb\b", &family)
            return smartbombMixes[StrLower(family[1])]
        ; Match the family only in the weapon field, including faction and
        ; advanced variants, never in a player's name.
        if RegExMatch(weapon, "i)\b(Mjolnir|Inferno|Scourge|Nova)\b", &family)
            return missileMixes[StrLower(family[1])]
        ; Only exact drone/fighter names are safe as a source fallback. Player names,
        ; turret names and NPC names cannot tell us which ammo was fired.
        attacker := Trim(source[1])
        if this.profiles.Has(attacker) && this.profiles[attacker].craft
            return this.profiles[attacker].mix
        return 0
    }

    LogSecond(tick := A_TickCount) {
        if !this.HasOwnProp("clockTick")
            return 0
        return this.clockSecond + Floor(Max(0, tick - this.clockTick) / 1000)
    }

    AddBatch(lines, tick := A_TickCount) {
        hits := [], latest := 0
        for line in lines {
            hit := this.ParseHit(line)
            if !IsObject(hit)
                continue
            hits.Push(hit)
            latest := Max(latest, hit.second)
        }
        if !hits.Length
            return
        ; The reader starts at EOF, so these are newly appended records. Anchor
        ; to their server timestamps rather than assuming Windows UTC agrees
        ; with EVE. Elapsed ticks keep the window expiring while the log is quiet.
        if !this.HasOwnProp("clockTick") || latest > this.LogSecond(tick) {
            this.clockSecond := latest
            this.clockTick := tick
        }
        now := this.LogSecond(tick)
        for hit in hits
            this.AddHit(hit, now)
    }

    AddLine(line, now) {
        ; Explicit log-time input is also useful for offline replay and tests.
        hit := this.ParseHit(line)
        if IsObject(hit)
            this.AddHit(hit, now)
    }

    ParseHit(line) {
        if !InStr(line, "(combat)")
            return
        ; Match the numeric damage immediately after (combat), followed by its
        ; direction. Neuts, repairs, misses and third-party tackle do not match.
        if !RegExMatch(line, "i)^\x{FEFF}?\s*\[\s*(\d{4}\.\d{2}\.\d{2} \d{2}:\d{2}:\d{2})\s*\]\s+\(combat\)\s*(?:<[^>]*>)*([0-9]+(?:\.[0-9]+)?)(?:<[^>]*>|\s)*(from|to)(?=<|\s)", &hit)
            return
        incoming := hit[3] = "from"
        if !this.settings[incoming ? "incomingDPSEnabled" : "outgoingDPSEnabled"]
            return
        if hit[1] != this.lastTimestamp {
            try second := DateDiff(RegExReplace(hit[1], "\D"), "19700101000000", "Seconds")
            catch
                return
            this.lastTimestamp := hit[1]
            this.lastSecond := second
        }
        return {second: this.lastSecond, incoming: incoming, amount: hit[2] + 0, line: line}
    }

    AddHit(hit, now) {
        second := hit.second
        ; Keep timestamp spacing within batches; old records do not all become
        ; simultaneous hits when several seconds arrive in one read.
        if second <= now - this.WindowSeconds || second > now
            return
        if !this.buckets.Has(second)
            this.buckets[second] := {incoming: 0, outgoing: 0, types: [0, 0, 0, 0, 0]}
        bucket := this.buckets[second]
        if hit.incoming {
            bucket.incoming += hit.amount
            if this.HasOwnProp("profiles") {
                mix := this.DamageMix(hit.line)
                if IsObject(mix) {
                    for index, fraction in mix
                        bucket.types[index] += hit.amount * fraction
                } else
                    bucket.types[5] += hit.amount
            }
        } else
            bucket.outgoing += hit.amount
    }

    Rates(now) {
        incoming := 0, outgoing := 0, expired := [], types := [0, 0, 0, 0, 0]
        for second, bucket in this.buckets {
            if second <= now - this.WindowSeconds || second > now
                expired.Push(second)
            else {
                incoming += bucket.incoming
                outgoing += bucket.outgoing
                if this.HasOwnProp("profiles") {
                    for index, amount in bucket.types
                        types[index] += amount
                }
            }
        }
        for second in expired
            this.buckets.Delete(second)
        return {incoming: incoming / this.WindowSeconds, outgoing: outgoing / this.WindowSeconds, types: types, damage: incoming}
    }

    Text(now) {
        rates := this.Rates(now)
        text := ""
        this.damageText := ""
        if this.settings["incomingDPSEnabled"] && rates.incoming > 0 && rates.incoming >= this.settings["incomingDPSThreshold"] {
            text := "In: " Round(rates.incoming)
            if this.HasOwnProp("profiles") {
                labels := ["EM", "Th", "Ki", "Ex", "?"]
                for index, amount in rates.types {
                    percentage := Round(amount / rates.damage * 100)
                    if percentage > 0
                        this.damageText .= (this.damageText = "" ? "%" : "") " " labels[index] " " percentage
                }
            }
        }
        if this.settings["outgoingDPSEnabled"] && rates.outgoing > 0 && rates.outgoing >= this.settings["outgoingDPSThreshold"]
            text .= (text = "" ? "" : " | ") "Out: " Round(rates.outgoing)
        return text = "" ? "" : "DPS: " text
    }
}
