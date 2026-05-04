-- doc export <path>  → PDF
set destPath to {{path|json}}
set destFile to POSIX file destPath as string
tell application "Microsoft Word"
    tell active document
        save as file name destFile file format format PDF
    end tell
end tell
return "exported"
