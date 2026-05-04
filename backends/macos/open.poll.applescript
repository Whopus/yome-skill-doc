tell application "Microsoft Word"
    if (count of documents) > 0 then
        return name of active document
    end if
    return ""
end tell
