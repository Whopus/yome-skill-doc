-- doc select <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        select text object of paragraph idx
    end tell
end tell
return "selected"
