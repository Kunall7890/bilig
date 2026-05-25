export function sheetDeletePromptHandlerAppleScript(): string {
  const command = `/usr/bin/osascript <<'BILIG_SHEET_DELETE_PROMPT' >/dev/null 2>&1 &
repeat 120 times
  delay 0.25
  tell application "System Events"
    tell process "Microsoft Excel"
      if exists button "Delete" of window 1 then
        click button "Delete" of window 1
        return
      end if
    end tell
  end tell
end repeat
BILIG_SHEET_DELETE_PROMPT`

  return `on startSheetDeletePromptHandler()
  do shell script ${toAppleScriptString(command)}
end startSheetDeletePromptHandler`
}

export function toAppleScriptString(value: string): string {
  return `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`
}

export function toAppleScriptValue(value: string | number | boolean): string {
  if (typeof value === 'string') {
    return toAppleScriptString(value)
  }
  if (typeof value === 'boolean') {
    return value ? 'true' : 'false'
  }
  return String(value)
}
