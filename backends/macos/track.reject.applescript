-- doc track.reject — reject all revisions
tell application "Microsoft Word"
    tell active document
        try
            reject all revisions
        on error
            try
                set revs to every revision
                repeat with r in revs
                    reject r
                end repeat
            end try
        end try
    end tell
end tell
return "rejected"
