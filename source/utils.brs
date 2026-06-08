' FormatTime: converts total seconds to "1h 23m" or "45m"
function FormatTime(seconds as Integer) as String
    if seconds <= 0 then return "0m"
    hours = int(seconds / 3600)
    mins = int((seconds mod 3600) / 60)
    if hours > 0
        return str(hours).trim() + "h " + str(mins).trim() + "m"
    else
        return str(mins).trim() + "m"
    end if
end function

' RelativeTime: converts ISO 8601 date string to relative time like "2 days ago"
function RelativeTime(dateStr as String) as String
    if dateStr = invalid or dateStr = "" then return ""

    now = CreateObject("roDateTime")
    now.Mark()

    dt = CreateObject("roDateTime")
    dt.FromISO8601String(dateStr)

    diffSeconds = now.AsSeconds() - dt.AsSeconds()

    if diffSeconds < 0 then return "just now"
    if diffSeconds < 60 then return "just now"
    if diffSeconds < 3600
        mins = int(diffSeconds / 60)
        if mins = 1 then return "1 minute ago"
        return str(mins).trim() + " minutes ago"
    end if
    if diffSeconds < 86400
        hrs = int(diffSeconds / 3600)
        if hrs = 1 then return "1 hour ago"
        return str(hrs).trim() + " hours ago"
    end if
    if diffSeconds < 2592000
        days = int(diffSeconds / 86400)
        if days = 1 then return "1 day ago"
        return str(days).trim() + " days ago"
    end if
    if diffSeconds < 31536000
        months = int(diffSeconds / 2592000)
        if months = 1 then return "1 month ago"
        return str(months).trim() + " months ago"
    end if
    years = int(diffSeconds / 31536000)
    if years = 1 then return "1 year ago"
    return str(years).trim() + " years ago"
end function
