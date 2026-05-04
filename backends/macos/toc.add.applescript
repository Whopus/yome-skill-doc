-- doc toc.add [--levels=N]
set lvl to ({{levels|json}}) as integer
if lvl < 1 then set lvl to 3
if lvl > 9 then set lvl to 9

tell application "Microsoft Word"
    tell active document
        set anchor to start of content of text object
        make new table of contents at active document with properties {range:text from start of text object to start of text object, lower heading level:1, upper heading level:lvl}
    end tell
end tell
return "added"
