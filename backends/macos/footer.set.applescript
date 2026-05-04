-- doc footer.set <text>
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        set ftr to footer (header footer index primary) of section 1
        set content of text object of ftr to theText
    end tell
end tell
return "set"
