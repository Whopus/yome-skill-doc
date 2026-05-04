-- doc link.add <index> --url=<url> [--text=<display>]
set idx to ({{index|json}}) as integer
set urlStr to {{url|json}}
set txtStr to {{text|json}}

tell application "Microsoft Word"
    tell active document
        set anchor to text object of paragraph idx
        if txtStr is not "" then
            set content of anchor to txtStr
        end if
        make new hyperlink at active document with properties {hyperlink address:urlStr, text object:anchor}
    end tell
end tell
return "added"
