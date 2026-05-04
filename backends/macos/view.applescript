-- doc view <print|web|draft|outline>
set modeStr to {{mode|json}}
set vt to print view
if modeStr is "web"     then set vt to web view
if modeStr is "draft"   then set vt to normal view
if modeStr is "outline" then set vt to outline view
tell application "Microsoft Word"
    if (count of windows) > 0 then
        try
            set type of view of active window to vt
        end try
    end if
end tell
return "set " & modeStr
