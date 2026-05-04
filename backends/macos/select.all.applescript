-- doc select.all
tell application "Microsoft Word"
    tell active document
        select text object
    end tell
end tell
return "selected"
