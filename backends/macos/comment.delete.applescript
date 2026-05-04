-- doc comment.delete <comment_idx>
set idx to ({{comment|json}}) as integer
tell application "Microsoft Word"
    tell active document
        delete comment idx
    end tell
end tell
return "deleted"
