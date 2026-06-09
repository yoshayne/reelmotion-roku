sub init()
    m.storedToken = readSessionToken()
    m.userData = invalid
    m.subscriptionActive = false
    m.activationScreen = invalid
    m.homeScreen = invalid
    m.detailScreen = invalid
    m.settingsScreen = invalid
    m.playerScreen = invalid

    if m.storedToken <> invalid and m.storedToken <> ""
        print "MainScene: session token found, showing home screen"
        showHomeScreen()
    else
        print "MainScene: no session token found, showing activation screen"
        showActivationScreen()
    end if
end sub

function readSessionToken() as String
    sec = CreateObject("roRegistrySection", "reelmotion")
    if sec.Exists("session_token")
        token = sec.Read("session_token")
        if token <> invalid
            return token
        end if
    end if
    return ""
end function

sub clearScreenStack()
    children = m.top.getChildren(-1, 0)
    for each child in children
        m.top.removeChild(child)
    end for

    m.activationScreen = invalid
    m.homeScreen = invalid
    m.detailScreen = invalid
    m.settingsScreen = invalid
    m.playerScreen = invalid
end sub

sub showActivationScreen()
    clearScreenStack()

    activation = CreateObject("roSGNode", "ActivationScreen")
    activation.observeField("activationComplete", "onActivationComplete")
    activation.observeField("close", "onActivationClose")
    m.top.appendChild(activation)
    activation.visible = true
    activation.setFocus(true)
    m.activationScreen = activation

    print "MainScene: showing activation screen"
end sub

sub onActivationComplete()
    if m.activationScreen = invalid then return
    if m.activationScreen.activationComplete <> true then return

    token = m.activationScreen.sessionToken
    if token = invalid or token = ""
        token = readSessionToken()
    end if

    if token = invalid or token = ""
        print "MainScene: activationComplete fired without a session token"
        return
    end if

    print "MainScene: activation complete, navigating to home"
    m.storedToken = token
    showHomeScreen()
end sub

sub onActivationClose()
    ' ActivationScreen currently does not emit close, but keep the observer
    ' available for future activation cancellation handling.
end sub

sub showHomeScreen()
    print "MainScene: showing home screen"
    clearScreenStack()

    home = CreateObject("roSGNode", "HomeScreen")
    if m.storedToken <> invalid and m.storedToken <> ""
        home.authToken = m.storedToken
    end if
    home.observeField("selectedItem", "onHomeItemSelected")
    home.observeField("goSettings", "onHomeGoSettings")
    home.observeField("close", "onHomeClose")
    m.top.appendChild(home)
    home.visible = true
    home.setFocus(true)
    m.homeScreen = home
end sub

sub onHomeClose()
    print "MainScene: home closed"
end sub

sub onHomeItemSelected()
    if m.homeScreen = invalid then return
    item = m.homeScreen.selectedItem
    if item = invalid then return
    showDetailScreen(item)
end sub

sub onHomeGoSettings()
    if m.homeScreen = invalid then return
    if m.homeScreen.goSettings = true
        showSettingsScreen()
    end if
end sub

sub showDetailScreen(item as Object)
    detail = CreateObject("roSGNode", "DetailScreen")
    if m.storedToken <> invalid and m.storedToken <> ""
        detail.authToken = m.storedToken
    end if
    detail.subscriptionActive = (m.subscriptionActive = true)
    detail.observeField("close", "onDetailClose")
    detail.observeField("playRequested", "onDetailPlayRequested")
    m.top.appendChild(detail)
    detail.visible = true
    detail.setFocus(true)
    m.detailScreen = detail
    if item.DoesExist("id")
        detail.contentId = item.id
    end if
end sub

sub onDetailClose()
    if m.detailScreen <> invalid
        m.top.removeChild(m.detailScreen)
        m.detailScreen = invalid
    end if
    if m.homeScreen <> invalid
        m.homeScreen.visible = true
        m.homeScreen.setFocus(true)
    end if
end sub

sub onDetailPlayRequested()
    if m.detailScreen = invalid then return
    if m.detailScreen.playRequested = true
        videoData = m.detailScreen.videoData
        showPlayerScreen(videoData)
    end if
end sub

sub showPlayerScreen(videoData as Object)
    player = CreateObject("roSGNode", "PlayerScreen")
    player.observeField("close", "onPlayerClose")
    m.top.appendChild(player)
    player.visible = true
    player.setFocus(true)
    m.playerScreen = player
    player.videoData = videoData
end sub

sub onPlayerClose()
    if m.playerScreen <> invalid
        m.top.removeChild(m.playerScreen)
        m.playerScreen = invalid
    end if
    if m.detailScreen <> invalid
        m.detailScreen.visible = true
        m.detailScreen.setFocus(true)
    else if m.homeScreen <> invalid
        m.homeScreen.setFocus(true)
    end if
end sub

sub showSettingsScreen()
    settings = CreateObject("roSGNode", "SettingsScreen")
    settings.subscriptionActive = (m.subscriptionActive = true)
    if m.userData <> invalid
        settings.userData = m.userData
    end if
    settings.observeField("close", "onSettingsClose")
    settings.observeField("signedOut", "onSignedOut")
    m.top.appendChild(settings)
    settings.visible = true
    settings.setFocus(true)
    m.settingsScreen = settings
end sub

sub onSettingsClose()
    if m.settingsScreen <> invalid
        m.top.removeChild(m.settingsScreen)
        m.settingsScreen = invalid
    end if
    if m.homeScreen <> invalid
        m.homeScreen.setFocus(true)
    end if
end sub

sub onSignedOut()
    print "MainScene: signed out, clearing registry and returning to activation"
    sec = CreateObject("roRegistrySection", "reelmotion")
    sec.Delete("session_token")
    sec.Flush()

    m.storedToken = ""
    m.userData = invalid
    m.subscriptionActive = false
    showActivationScreen()
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press and key = "back"
        if m.playerScreen <> invalid
            onPlayerClose()
            return true
        end if
        if m.detailScreen <> invalid
            onDetailClose()
            return true
        end if
        if m.settingsScreen <> invalid
            onSettingsClose()
            return true
        end if
    end if
    return false
end function
