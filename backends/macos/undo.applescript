-- doc undo
tell application "Microsoft Word"
    try
        undo
    on error
        try
            tell active document to undo
        end try
    end try
end tell
return "undone"
