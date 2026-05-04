-- doc list.outdent <index>
set idx to ({{index|json}}) as integer
tell application "Microsoft Word"
    tell active document
        set pf to paragraph format of text object of paragraph idx
        set newIndent to (left indent of pf) - 36
        if newIndent < 0 then set newIndent to 0
        set left indent of pf to newIndent
    end tell
end tell
return "outdented"
