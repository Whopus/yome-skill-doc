-- doc cursor.end — collapse selection to end of document
tell application "Microsoft Word"
    tell active document
        set sel to selection
        set selEnd to end of content of text object
        set selection start of sel to selEnd
        set selection end of sel to selEnd
    end tell
end tell
return "moved"
