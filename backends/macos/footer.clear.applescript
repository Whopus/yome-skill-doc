-- doc footer.clear
tell application "Microsoft Word"
    tell active document
        set ftr to footer (header footer index primary) of section 1
        set content of text object of ftr to ""
    end tell
end tell
return "cleared"
