-- doc alerts <on|off>
set spec to {{state|json}}
set flag to true
if spec is "off" or spec is "false" then set flag to false
tell application "Microsoft Word"
    set display alerts to flag
end tell
if flag then
    return "on"
else
    return "off"
end if
