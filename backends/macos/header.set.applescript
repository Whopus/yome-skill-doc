-- doc header.set <text>
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        set hdr to header (header footer index primary) of section 1
        set content of text object of hdr to theText
    end tell
end tell
return "set"
