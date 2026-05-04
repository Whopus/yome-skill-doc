-- doc table.col.delete <table_idx> --col=N
set tblIdx to ({{table|json}}) as integer
set colIdx to ({{col|json}}) as integer
tell application "Microsoft Word"
    tell active document
        tell table tblIdx
            delete column colIdx
        end tell
    end tell
end tell
return "deleted"
