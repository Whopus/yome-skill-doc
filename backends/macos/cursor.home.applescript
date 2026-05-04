-- doc cursor.home — collapse selection to start of document
tell application "Microsoft Word"
    tell active document
        set sel to selection
        set selStart to start of content of text object
        set selection start of sel to selStart
        set selection end of sel to selStart
    end tell
end tell
return "moved"
