-- doc comment.add <index> --text=<content>
set idx to ({{index|json}}) as integer
set theText to {{text|json}}
tell application "Microsoft Word"
    tell active document
        set anchor to text object of paragraph idx
        make new comment at active document with properties {comment text:theText, scope:anchor}
    end tell
end tell
return "added"
