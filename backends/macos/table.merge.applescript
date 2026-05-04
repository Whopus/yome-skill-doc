-- doc table.merge <table_idx> --fromRow=N --fromCol=N --toRow=N --toCol=N
set tblIdx to ({{table|json}}) as integer
set fr to ({{fromRow|json}}) as integer
set fc to ({{fromCol|json}}) as integer
set tr to ({{toRow|json}}) as integer
set tc to ({{toCol|json}}) as integer

tell application "Microsoft Word"
    tell active document
        tell table tblIdx
            set c1 to cell from row fr column fc
            set c2 to cell from row tr column tc
            merge c1 with c2
        end tell
    end tell
end tell
return "merged"
