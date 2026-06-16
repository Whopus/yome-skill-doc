-- doc save [<path>] [--path=<save_as>] [--force]
set savePath to {{path|json}}

if savePath is "" then
    tell application "Microsoft Word"
        tell active document
            save
        end tell
    end tell
    return "saved"
end if

set destPath to POSIX file savePath as string
set ext to ""
set AppleScript's text item delimiters to "."
set parts to text items of savePath
set AppleScript's text item delimiters to ""
if (count of parts) > 1 then set ext to last item of parts

tell application "Microsoft Word"
    try
        if ext is "doc" then
            save as document 1 file name destPath file format format document97
        else
            save as document 1 file name destPath file format format document default
        end if
    on error errMsg
        return "ERROR: save failed at " & savePath & " — " & errMsg
    end try
end tell
return "saved"
