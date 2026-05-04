-- doc table.add [--index=<para_idx>] [--rows=N] [--cols=N]
set idxStr to {{index|json}}
set numRows to ({{rows|json}}) as integer
set numCols to ({{cols|json}}) as integer
if numRows < 1 then set numRows to 3
if numCols < 1 then set numCols to 3

tell application "Microsoft Word"
    tell active document
        if idxStr is "" then
            set anchor to text object
        else
            set anchor to text object of paragraph ((idxStr as integer))
        end if
        make new table at anchor with properties {number of rows:numRows, number of columns:numCols}
        return (count of tables) as string
    end tell
end tell
