-- doc track.accept — accept all revisions
tell application "Microsoft Word"
    tell active document
        try
            accept all revisions
        on error
            try
                set revs to every revision
                repeat with r in revs
                    accept r
                end repeat
            end try
        end try
    end tell
end tell
return "accepted"
