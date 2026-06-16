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

-- Pick file format from extension: .docx → XML document, .doc → Word document
set ext to ""
set AppleScript's text item delimiters to "."
set parts to text items of savePath
set AppleScript's text item delimiters to ""
if (count of parts) > 1 then set ext to last item of parts

set formatClause to "format XML document"
if ext is "doc" then set formatClause to "format document"

tell application "Microsoft Word"
    tell active document
        try
            save as file name destPath file format format XML document
        on error
            -- Some Word versions only accept the bare format token
            try
                save as file name destPath
            on error errMsg
                return "ERROR: save failed at " & savePath & " — " & errMsg
            end try
        end try
    end tell
end tell
return "saved"
