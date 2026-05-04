-- doc header.clear
tell application "Microsoft Word"
    tell active document
        set hdr to header (header footer index primary) of section 1
        set content of text object of hdr to ""
    end tell
end tell
return "cleared"
