-- doc table.row.add <table_idx> [--count=1]
set tblIdx to ({{table|json}}) as integer
set n to ({{count|json}}) as integer
if n < 1 then set n to 1

tell application "Microsoft Word"
    tell active document
        repeat n times
            tell table tblIdx
                make new row at end
            end tell
        end repeat
    end tell
end tell
return "added " & (n as string)
