-- doc table.row.delete <table_idx> --row=N
set tblIdx to ({{table|json}}) as integer
set rowIdx to ({{row|json}}) as integer
tell application "Microsoft Word"
    tell active document
        tell table tblIdx
            delete row rowIdx
        end tell
    end tell
end tell
return "deleted"
