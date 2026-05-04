-- doc redo
tell application "Microsoft Word"
    try
        redo
    on error
        try
            tell active document to redo
        end try
    end try
end tell
return "redone"
