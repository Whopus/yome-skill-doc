-- doc list.bullet <index>  — apply bullet list format
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        set p to paragraph idx
        try
            apply list format hanging of list format of text object of p list type bulleted
        on error
            -- Fall back to applying the built-in "List Bullet" style
            try
                set style of p to style "List Bullet" of active document
            end try
        end try
    end tell
end tell
return "applied"
