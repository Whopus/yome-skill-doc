-- doc set <index> --text=<content>
set idx to ({{index|json}}) as integer
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        set content of text object of paragraph idx to theText & return
    end tell
end tell
return "updated"
