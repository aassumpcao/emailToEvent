-- Created by Andre Assumpcao
-- andre.assumpcao@gmail.com

on run {input, parameters}

  tell application "Microsoft Outlook"
    repeat with theMessage in (get selection)
      set theContent to content of theMessage
      set theSubject to subject of theMessage
    end repeat
  end tell

  set theDate to current date
  set theText to the text returned of (display dialog "Date and Time:" default answer (short date string of theDate & space & time string of theDate))
  set theStartDate to date theText
  set theEndDate to theStartDate + (1 * hours)
  set theLocation to the text returned of (display dialog "Location:" default answer "")

  tell application "Calendar"
    set CalendarList to name of every calendar
    set theCalendar to (choose from list CalendarList with title "Choose a Calendar")

    tell calendar (get theCalendar as text)
      make new event with properties {summary:theSubject, start date:theStartDate, end date:theEndDate, location:theLocation, description:theContent}
    end tell
  end tell

  return input

end run