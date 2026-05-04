-- doc list.indent <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        set pf to paragraph format of text object of paragraph idx
        set left indent of pf to (left indent of pf) + 36
    end tell
end tell
return "indented"
