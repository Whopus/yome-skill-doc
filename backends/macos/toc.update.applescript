-- doc toc.update — refresh all TOCs
tell application "Microsoft Word"
    tell active document
        set tocs to every table of contents
        repeat with t in tocs
            try
                update t
            end try
        end repeat
    end tell
end tell
return "updated"
