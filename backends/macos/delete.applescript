-- doc delete <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        delete text object of paragraph idx
    end tell
end tell
return "deleted"
