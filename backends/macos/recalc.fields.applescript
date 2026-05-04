-- doc recalc.fields — update every field in the document
tell application "Microsoft Word"
    tell active document
        try
            update fields
        on error
            try
                set flds to every field
                repeat with f in flds
                    try
                        update f
                    end try
                end repeat
            end try
        end try
    end tell
end tell
return "updated"
