-- doc insert <index> --text=<content>
set idx to ({{index|json}}) as integer
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        insert text (theText & return) at beginning of text object of paragraph idx
    end tell
end tell
return "inserted"
