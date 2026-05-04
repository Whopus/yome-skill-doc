-- doc list.number <index>  — apply numbered list format
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        set p to paragraph idx
        try
            apply list format hanging of list format of text object of p list type numbered
        on error
            try
                set style of p to style "List Number" of active document
            end try
        end try
    end tell
end tell
return "applied"
