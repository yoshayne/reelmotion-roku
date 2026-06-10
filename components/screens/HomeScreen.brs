sub init()
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onContentLoaded")
    m.httpTask.control = "RUN"

    m.top.observeField("authToken", "onAuthToken")

    m.rowHeight = 320
    m.colWidth = 206
    m.focusRow = 0
    m.focusCol = 0
    m.cards = []
end sub

sub onAuthToken()
    loadContent()
end sub

sub loadContent()
    token = m.top.authToken
    headers = {}
    if token <> invalid and token <> ""
        headers = {Authorization: "Bearer " + token}
    end if

    m.top.findNode("spinner").visible = true
    m.top.findNode("errorLabel").visible = false

    m.httpTask.request = {
        url: "https://reelmotionapp.com/api/browse-data",
        method: "GET",
        headers: headers,
        body: "",
        context: "browseData"
    }
end sub

sub onContentLoaded()
    resp = m.httpTask.response
    m.top.findNode("spinner").visible = false

    if resp = invalid or resp.code <> 200
        m.top.findNode("errorLabel").text = "Could not load content. Please try again."
        m.top.findNode("errorLabel").visible = true
        return
    end if

    json = ParseJson(resp.content)
    if json = invalid
        m.top.findNode("errorLabel").text = "Error parsing content."
        m.top.findNode("errorLabel").visible = true
        return
    end if

    buildGrid(json)
end sub

sub buildGrid(json as Object)
    container = m.top.findNode("gridContainer")
    container.removeChildren(container.getChildren(-1, 0))
    m.cards = []

    rows = []

    if json.continue_watching <> invalid and json.continue_watching.count() > 0
        rows.push({ title: "Continue Watching", videos: json.continue_watching })
    end if

    if json.categories <> invalid
        for each cat in json.categories
            videos = []
            if cat.videos <> invalid then videos = cat.videos
            rows.push({ title: cat.name, videos: videos })
        end for
    end if

    yOffset = 0
    for r = 0 to rows.count() - 1
        row = rows[r]

        titleLabel = CreateObject("roSGNode", "Label")
        titleLabel.text = row.title
        titleLabel.font = "font:MediumBoldSystemFont"
        titleLabel.color = "#FFFFFF"
        titleLabel.translation = [0, yOffset]
        container.appendChild(titleLabel)

        rowCards = []
        for c = 0 to row.videos.count() - 1
            video = row.videos[c]
            card = CreateObject("roSGNode", "PosterCard")
            card.translation = [c * m.colWidth, yOffset + 40]
            card.itemContent = makeVideoNode(video)
            container.appendChild(card)
            rowCards.push(card)
        end for

        m.cards.push(rowCards)
        yOffset = yOffset + m.rowHeight
    end for

    if m.cards.count() > 0 and m.cards[0].count() > 0
        m.cards[0][0].isFocused = true
        m.top.setFocus(true)
    end if
end sub

function makeVideoNode(video as Object) as Object
    item = CreateObject("roSGNode", "ContentNode")

    title = video.title
    if title = invalid then title = ""
    item.title = title

    thumbnail = video.thumbnail
    if thumbnail = invalid then thumbnail = ""
    if Left(thumbnail, 1) = "/"
        thumbnail = "https://reelmotionapp.com" + thumbnail
    end if
    item.HDPosterUrl = thumbnail

    rating = video.rating
    if rating = invalid then rating = ""
    item.shortDescriptionLine1 = rating

    duration = video.duration
    if duration = invalid then duration = ""
    item.shortDescriptionLine2 = duration

    ' Store custom fields
    item.addFields({
        id: "",
        is_free: false
    })
    if video.id <> invalid
        item.id = video.id.toStr()
    end if
    if video.is_free <> invalid
        item.is_free = video.is_free
    end if

    return item
end function

sub moveFocus(newRow as Integer, newCol as Integer)
    if newRow < 0 or newRow >= m.cards.count() then return
    if newCol < 0 or newCol >= m.cards[newRow].count() then return

    m.cards[m.focusRow][m.focusCol].isFocused = false
    m.focusRow = newRow
    m.focusCol = newCol
    m.cards[m.focusRow][m.focusCol].isFocused = true

    container = m.top.findNode("gridContainer")
    targetY = 90 - (m.focusRow * m.rowHeight)
    if targetY > 90 then targetY = 90
    container.translation = [60, targetY]
end sub

sub onItemSelected()
    if m.cards.count() = 0 then return
    selected = m.cards[m.focusRow][m.focusCol]
    m.top.selectedItem = selected.itemContent
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            m.top.close = true
            return true
        end if
        if m.cards.count() = 0 then return false
        if key = "up"
            moveFocus(m.focusRow - 1, m.focusCol)
            return true
        end if
        if key = "down"
            moveFocus(m.focusRow + 1, m.focusCol)
            return true
        end if
        if key = "left"
            moveFocus(m.focusRow, m.focusCol - 1)
            return true
        end if
        if key = "right"
            moveFocus(m.focusRow, m.focusCol + 1)
            return true
        end if
        if key = "OK"
            onItemSelected()
            return true
        end if
        if key = "options"
            m.top.goSettings = true
            return true
        end if
    end if
    return false
end function
