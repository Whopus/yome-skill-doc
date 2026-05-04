-- doc track.on
tell application "Microsoft Word"
    tell active document
        set track revisions to true
    end tell
end tell
return "on"
