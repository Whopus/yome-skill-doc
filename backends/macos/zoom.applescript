-- doc zoom <level>  (10..400)
set lvl to ({{level|json}}) as integer
if lvl < 10  then set lvl to 10
if lvl > 400 then set lvl to 400
tell application "Microsoft Word"
    if (count of windows) > 0 then
        try
            set percentage of zoom of view of active window to lvl
        end try
    end if
end tell
return "set " & (lvl as string)
