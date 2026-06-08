sub init()
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onHttpResponse")

    m.video = m.top.findNode("video")
    m.video.observeField("state", "onVideoState")
    m.video.observeField("position", "onVideoPosition")

    m.top.findNode("skipIntroBtn").observeField("buttonSelected", "onSkipIntro")
    m.top.findNode("retryBtn").observeField("buttonSelected", "onRetry")

    m.introEndSeconds = 0
    m.showingIntro = false
    m.videoId = ""
    m.lastPosition = 0

    m.top.observeField("videoData", "onVideoData")
end sub

sub onVideoData()
    data = m.top.videoData
    if data = invalid then return

    muxId = data.mux_playback_id
    if muxId = invalid or muxId = ""
        showError("Invalid video source.")
        return
    end if

    m.videoId = ""
    if data.id <> invalid
        if type(data.id) = "String"
            m.videoId = data.id
        else
            m.videoId = data.id.toStr()
        end if
    end if

    m.introEndSeconds = 0
    if data.intro_end_seconds <> invalid
        m.introEndSeconds = data.intro_end_seconds
        m.showingIntro = true
    end if

    showLoadingOverlay(true)

    hlsUrl = "https://stream.mux.com/" + muxId + ".m3u8"

    content = CreateObject("roSGNode", "ContentNode")
    content.url = hlsUrl
    content.streamFormat = "hls"
    content.title = data.title

    m.video.content = content
    m.video.control = "play"
end sub

sub onVideoState()
    state = m.video.state
    if state = "playing"
        showLoadingOverlay(false)
        hideError()
    else if state = "buffering"
        showLoadingOverlay(true)
    else if state = "stopped"
        ' Video ended or stopped
    else if state = "error"
        showLoadingOverlay(false)
        showError("Playback error. Please try again.")
    end if
end sub

sub onVideoPosition()
    pos = m.video.position
    m.lastPosition = pos

    if m.showingIntro and m.introEndSeconds > 0
        if pos < m.introEndSeconds
            m.top.findNode("skipIntroBtn").visible = true
        else
            m.top.findNode("skipIntroBtn").visible = false
            m.showingIntro = false
        end if
    end if
end sub

sub onSkipIntro()
    m.video.seek = m.introEndSeconds
    m.top.findNode("skipIntroBtn").visible = false
    m.showingIntro = false
end sub

sub onRetry()
    hideError()
    onVideoData()
end sub

sub showLoadingOverlay(show as Boolean)
    m.top.findNode("loadingBg").visible = show
    m.top.findNode("loadingLogo").visible = show
    m.top.findNode("loadingSpinner").visible = show
end sub

sub showError(msg as String)
    showLoadingOverlay(false)
    m.top.findNode("errorBg").visible = true
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("retryBtn").visible = true
    m.top.findNode("retryBtn").setFocus(true)
end sub

sub hideError()
    m.top.findNode("errorBg").visible = false
    m.top.findNode("errorLabel").visible = false
    m.top.findNode("retryBtn").visible = false
end sub

sub savePlaybackHistory()
    if m.videoId = "" or m.lastPosition <= 0 then return

    token = m.top.authToken
    headers = {"Content-Type": "application/json"}
    if token <> invalid and token <> ""
        headers["Authorization"] = "Bearer " + token
    end if

    body = FormatJson({video_id: m.videoId, position: m.lastPosition})

    m.httpTask.request = {
        url: "https://www.reelmotionapp.com/api/playback-history",
        method: "POST",
        headers: headers,
        body: body,
        context: "playbackHistory"
    }
end sub

sub onHttpResponse()
    ' Playback history response - nothing to do on success
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            m.video.control = "stop"
            savePlaybackHistory()
            m.top.close = true
            return true
        end if
        if key = "OK"
            skipBtn = m.top.findNode("skipIntroBtn")
            if skipBtn.visible = true and skipBtn.hasFocus()
                onSkipIntro()
                return true
            end if
        end if
    end if
    return false
end function
