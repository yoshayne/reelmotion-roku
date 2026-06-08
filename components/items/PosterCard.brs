sub init()
    m.top.observeField("itemContent", "onItemContent")
    m.top.observeField("focusPercent", "onFocusPercent")
end sub

sub onItemContent()
    content = m.top.itemContent
    if content = invalid then return

    poster = m.top.findNode("poster")
    if content.HDPosterUrl <> invalid and content.HDPosterUrl <> ""
        poster.uri = content.HDPosterUrl
    end if

    titleLabel = m.top.findNode("titleLabel")
    if content.title <> invalid
        titleLabel.text = content.title
    end if
end sub

sub onFocusPercent()
    focusBorder = m.top.findNode("focusBorder")
    pct = m.top.focusPercent
    if pct > 0.5
        focusBorder.opacity = 1.0
    else
        focusBorder.opacity = 0.0
    end if
end sub
