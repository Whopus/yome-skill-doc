-- doc columns <count>
set n to ({{count|json}}) as integer
if n < 1 then set n to 1
if n > 6 then set n to 6
tell application "Microsoft Word"
    tell active document
        try
            set text columns of page setup of section 1 to n
        end try
    end tell
end tell
return "set " & (n as string)
