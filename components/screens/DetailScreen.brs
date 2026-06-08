sub init()
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onHttpResponse")

    m.videoLoaded = false
    m.commentsLoaded = false
    m.videoData = invalid
    m.commentsData = []

    m.top.findNode("playBtn").observeField("buttonSelected", "onPlaySelected")
    m.top.findNode("myListBtn").observeField("buttonSelected", "onMyListSelected")
    m.top.findNode("moreInfoBtn").observeField("buttonSelected", "onMoreInfoSelected")

    m.top.observeField("contentId", "onContentId")
    m.top.observeField("subscriptionActive", "onSubscriptionChanged")
end sub

sub onContentId()
    contentId = m.top.contentId
    if contentId = invalid or contentId = "" then return
    loadData(contentId)
end sub

sub loadData(contentId as String)
    m.videoLoaded = false
    m.commentsLoaded = false
    m.top.findNode("spinner").visible = true

    token = m.top.authToken
    headers = {}
    if token <> invalid and token <> ""
        headers = {Authorization: "Bearer " + token}
    end if

    m.pendingVideoId = contentId

    m.httpTask.request = {
        url: "https://www.reelmotionapp.com/api/watch/" + contentId,
        method: "GET",
        headers: headers,
        body: "",
        context: "videoDetail"
    }
end sub

sub onHttpResponse()
    resp = m.httpTask.response
    if resp = invalid then return

    context = resp.context

    if context = "videoDetail"
        m.videoLoaded = true
        if resp.code = 200
            json = ParseJson(resp.content)
            if json <> invalid
                m.videoData = json
                m.top.videoData = json
                populateVideoUI(json)
            end if
        end if

        ' Now load comments
        token = m.top.authToken
        headers = {}
        if token <> invalid and token <> ""
            headers = {Authorization: "Bearer " + token}
        end if
        m.httpTask.request = {
            url: "https://www.reelmotionapp.com/api/videos/" + m.pendingVideoId + "/comments",
            method: "GET",
            headers: headers,
            body: "",
            context: "comments"
        }

    else if context = "comments"
        m.commentsLoaded = true
        m.top.findNode("spinner").visible = false
        if resp.code = 200
            json = ParseJson(resp.content)
            if json <> invalid
                m.commentsData = json
                populateCommentsUI(json)
            end if
        end if
        ' Set initial focus
        updateButtonVisibility()
        if m.top.findNode("playBtn").visible = true
            m.top.findNode("playBtn").setFocus(true)
        else
            m.top.findNode("myListBtn").setFocus(true)
        end if
    end if
end sub

sub populateVideoUI(data as Object)
    if data.title <> invalid
        m.top.findNode("titleLabel").text = data.title
    end if

    meta = ""
    if data.year <> invalid then meta = meta + str(data.year).trim()
    if data.rating <> invalid and data.rating <> ""
        if meta <> "" then meta = meta + "  |  "
        meta = meta + data.rating
    end if
    if data.duration <> invalid
        if meta <> "" then meta = meta + "  |  "
        durationStr = ""
        if type(data.duration) = "Integer" or type(data.duration) = "roInt"
            durationStr = FormatTime(data.duration)
        else
            durationStr = data.duration.toStr()
        end if
        meta = meta + durationStr
    end if
    if data.director <> invalid and data.director <> ""
        if meta <> "" then meta = meta + "  |  "
        meta = meta + "Dir: " + data.director
    end if
    m.top.findNode("metaLabel").text = meta

    if data.description <> invalid
        m.top.findNode("descLabel").text = data.description
    end if

    if data.thumbnail <> invalid
        m.top.findNode("heroPoster").uri = data.thumbnail
    end if

    updateButtonVisibility()
end sub

sub updateButtonVisibility()
    if m.videoData = invalid then return

    isFree = false
    if m.videoData.is_free <> invalid then isFree = (m.videoData.is_free = true)
    subscriptionActive = (m.top.subscriptionActive = true)

    canPlay = isFree or subscriptionActive
    m.top.findNode("playBtn").visible = canPlay
    m.top.findNode("membersLabel").visible = not canPlay
end sub

sub onSubscriptionChanged()
    updateButtonVisibility()
end sub

sub populateCommentsUI(comments as Object)
    if comments = invalid or comments.count() = 0
        m.top.findNode("commentsLabel").text = "No comments yet."
        return
    end if

    commentText = ""
    maxComments = 3
    count = 0
    for each c in comments
        if count >= maxComments then exit for
        if c.display_name <> invalid
            commentText = commentText + c.display_name + ": "
        end if
        if c.body <> invalid
            commentText = commentText + c.body
        end if
        if c.created_at <> invalid
            commentText = commentText + "  (" + RelativeTime(c.created_at) + ")"
        end if
        commentText = commentText + Chr(10)
        count = count + 1
    end for
    m.top.findNode("commentsLabel").text = commentText
end sub

sub onPlaySelected()
    m.top.playRequested = true
end sub

sub onMyListSelected()
    ' My List functionality placeholder - no backend endpoint defined
end sub

sub onMoreInfoSelected()
    ' Could expand description or show additional details
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            m.top.close = true
            return true
        end if
    end if
    return false
end function

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

function RelativeTime(dateStr as String) as String
    if dateStr = invalid or dateStr = "" then return ""
    now = CreateObject("roDateTime")
    now.Mark()
    dt = CreateObject("roDateTime")
    dt.FromISO8601String(dateStr)
    diffSeconds = now.AsSeconds() - dt.AsSeconds()
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
    days = int(diffSeconds / 86400)
    if days = 1 then return "1 day ago"
    return str(days).trim() + " days ago"
end function
