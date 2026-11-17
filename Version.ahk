; Single source for script, executable and Settings editor release versions.
global ReleaseVersion := "1.7.0.8"
;@Ahk2Exe-Let U_version = %A_PriorLine~^.*"([0-9.]+)".*$~$1%
;@Ahk2Exe-SetVersion %U_version%
;@Ahk2Exe-SetFileVersion %U_version%
