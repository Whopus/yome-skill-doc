-- doc print [--copies=N]
set n to ({{copies|json}}) as integer
if n < 1 then set n to 1
tell application "Microsoft Word"
    tell active document
        print copies n
    end tell
end tell
return "printed"
