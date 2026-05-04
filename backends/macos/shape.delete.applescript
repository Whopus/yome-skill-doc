-- doc shape.delete <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        delete shape idx
    end tell
end tell
return "deleted"
