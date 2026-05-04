-- doc pagenum [--where=header|footer] [--align=left|center|right]
set whereStr to {{where|json}}
if whereStr is "" then set whereStr to "footer"
set alignStr to {{align|json}}
if alignStr is "" then set alignStr to "center"

set alignVal to align paragraph center
if alignStr is "left"  then set alignVal to align paragraph left
if alignStr is "right" then set alignVal to align paragraph right

tell application "Microsoft Word"
    tell active document
        if whereStr is "header" then
            set hf to header (header footer index primary) of section 1
        else
            set hf to footer (header footer index primary) of section 1
        end if
        tell hf
            set selRange to text object
            set paragraph alignment of paragraph format of selRange to alignVal
            make new field at end of selRange with properties {field type:field page}
        end tell
    end tell
end tell
return "added"
