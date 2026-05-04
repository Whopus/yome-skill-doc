-- doc toc.delete — remove all tables of contents
tell application "Microsoft Word"
    tell active document
        delete every table of contents
    end tell
end tell
return "deleted"
