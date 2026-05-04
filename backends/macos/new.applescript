-- doc new [<path>] [--force]
set savePath to {{path|json}}
set forceFlag to {{force|bool}}

tell application "Microsoft Word"
    activate
    delay 0.5
    set newDoc to make new document
    delay 0.3
end tell

if savePath is not "" then
    set destPath to POSIX file savePath as string
    tell application "Microsoft Word"
        try
            tell active document
                save as file name destPath file format format document
            end tell
        on error errMsg
            return "ERROR: save failed at " & savePath & " — " & errMsg
        end try
    end tell
    return "created and saved"
end if
return "created"
