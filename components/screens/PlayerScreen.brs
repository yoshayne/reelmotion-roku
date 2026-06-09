sub init()
    m.video = CreateObject("roSGNode", "Video")
    m.video.width = 1280
    m.video.height = 720
    m.top.appendChild(m.video)
    m.top.observeField("videoData", "onVideoData")
end sub

sub onVideoData()
    data = m.top.videoData
    if data = invalid then return
    if data.mux_playback_id = invalid then return

    url = "https://stream.mux.com/" + data.mux_playback_id + ".m3u8"

    content = CreateObject("roSGNode", "ContentNode")
    content.url = url
    content.streamFormat = "hls"
    if data.title <> invalid
        content.title = data.title
    end if

    m.video.content = content
    m.video.observeField("state", "onVideoStateChange")
    m.video.control = "play"
    m.video.setFocus(true)
end sub

sub onVideoStateChange()
    state = m.video.state
    if state = "finished" or state = "error"
        closePlayer()
    end if
end sub

sub closePlayer()
    if m.video <> invalid
        m.video.control = "stop"
        m.video.content = invalid
    end if
    m.top.close = true
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            closePlayer()
            return true
        end if
    end if
    return false
end function
