-- doc track.off
tell application "Microsoft Word"
    tell active document
        set track revisions to false
    end tell
end tell
return "off"
