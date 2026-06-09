sub init()
    m.screenStack = []
    m.currentScreen = invalid

    m.registryTask = CreateObject("roSGNode", "RegistryTask")
    m.registryTask.control = "RUN"

    m.top.observeField("launchArgs", "onLaunchArgs")

    CheckStoredToken()
end sub

sub onLaunchArgs()
    ' Handle deep link args if needed
    args = m.top.launchArgs
    if args <> invalid
        ' future: handle contentId deep links
    end if
end sub

sub CheckStoredToken()
    ' Use RegistryTask to read device_token
    context = CreateObject("roSGNode", "Node")
    context.addFields({
        parameters: {
            command: "read",
            section: "reelmotion",
            key: "device_token",
            value: ""
        },
        response: {}
    })
    context.observeField("response", "onTokenRead")
    m.registryTask.request = {context: context}
end sub

sub onTokenRead()
    context = m.registryTask.request.context
    if context = invalid then
        ShowActivationScreen()
        return
    end if
    resp = context.response
    token = ""
    if resp <> invalid and resp.DoesExist("regVal")
        token = resp.regVal
    end if

    if token = invalid or token = "" or token = "invalid"
        ShowActivationScreen()
    else
        m.storedToken = token
        VerifyToken(token)
    end if
end sub

sub VerifyToken(token as String)
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onVerifyResponse")
    m.httpTask.request = {
        url: "https://reelmotionapp.com/api/auth/device/verify",
        method: "GET",
        headers: {
            "Authorization": "Bearer " + token
        },
        body: "",
        context: "verify"
    }
end sub

sub onVerifyResponse()
    resp = m.httpTask.response
    if resp = invalid
        ShowActivationScreen()
        return
    end if

    if resp.code = 200
        json = ParseJson(resp.content)
        if json <> invalid
            m.userData = json.user
            m.subscriptionActive = (json.subscription <> invalid and json.subscription <> false)
            ShowHomeScreen()
        else
            ClearToken()
            ShowActivationScreen()
        end if
    else
        ClearToken()
        ShowActivationScreen()
    end if
end sub

sub ClearToken()
    context = CreateObject("roSGNode", "Node")
    context.addFields({
        parameters: {
            command: "delete",
            section: "reelmotion",
            key: "device_token",
            value: ""
        },
        response: {}
    })
    m.registryTask.request = {context: context}
    m.storedToken = ""
end sub

'--------------------------------------------------
' Screen Stack Management
'--------------------------------------------------

sub ShowScreen(screen as Object)
    if m.currentScreen <> invalid
        m.screenStack.push(m.currentScreen)
        m.currentScreen.visible = false
    end if
    m.currentScreen = screen
    m.top.findNode("mainSceneId").appendChild(screen)
    screen.visible = true
    screen.setFocus(true)
end sub

sub CloseScreen()
    if m.currentScreen <> invalid
        m.top.findNode("mainSceneId").removeChild(m.currentScreen)
        m.currentScreen = invalid
    end if
    if m.screenStack.count() > 0
        m.currentScreen = m.screenStack.pop()
        m.currentScreen.visible = true
        m.currentScreen.setFocus(true)
    end if
end sub

'--------------------------------------------------
' Screen Factories
'--------------------------------------------------

sub ShowActivationScreen()
    screen = CreateObject("roSGNode", "ActivationScreen")
    screen.observeField("goHome", "onActivationGoHome")
    ShowScreen(screen)
end sub

sub onActivationGoHome()
    if m.currentScreen = invalid then return
    goHome = m.currentScreen.goHome
    if goHome = true
        CloseScreen()
        ' re-read token and go home
        CheckStoredToken()
    end if
end sub

sub ShowHomeScreen()
    screen = CreateObject("roSGNode", "HomeScreen")
    if m.storedToken <> invalid
        screen.authToken = m.storedToken
    end if
    screen.observeField("selectedItem", "onHomeItemSelected")
    screen.observeField("goSettings", "onHomeGoSettings")
    screen.observeField("close", "onScreenClose")
    ShowScreen(screen)
end sub

sub onHomeItemSelected()
    if m.currentScreen = invalid then return
    item = m.currentScreen.selectedItem
    if item = invalid then return
    ShowDetailScreen(item)
end sub

sub onHomeGoSettings()
    if m.currentScreen = invalid then return
    goSettings = m.currentScreen.goSettings
    if goSettings = true
        ShowSettingsScreen()
    end if
end sub

sub ShowDetailScreen(item as Object)
    screen = CreateObject("roSGNode", "DetailScreen")
    if m.storedToken <> invalid
        screen.authToken = m.storedToken
    end if
    screen.subscriptionActive = (m.subscriptionActive = true)
    screen.observeField("close", "onScreenClose")
    screen.observeField("playRequested", "onDetailPlayRequested")
    m.detailScreen = screen
    ShowScreen(screen)
    ' Set contentId after show so observers are ready
    if item.DoesExist("id")
        screen.contentId = item.id
    end if
end sub

sub onDetailPlayRequested()
    if m.currentScreen = invalid then return
    play = m.currentScreen.playRequested
    if play = true
        videoData = m.currentScreen.videoData
        ShowPlayerScreen(videoData)
    end if
end sub

sub ShowPlayerScreen(videoData as Object)
    ' PlayerScreen extends Scene so it must be appended directly to m.top,
    ' not to the mainSceneId Rectangle like other screens.
    screen = CreateObject("roSGNode", "PlayerScreen")
    screen.observeField("close", "onPlayerClose")
    m.top.appendChild(screen)
    screen.visible = true
    screen.setFocus(true)
    m.playerScreen = screen
    screen.videoData = videoData
end sub

sub onPlayerClose()
    if m.playerScreen <> invalid
        m.playerScreen.visible = false
        m.top.removeChild(m.playerScreen)
        m.playerScreen = invalid
    end if
    prev = m.screenStack.Peek()
    if prev <> invalid
        prev.visible = true
        prev.setFocus(true)
    else if m.currentScreen <> invalid
        m.currentScreen.visible = true
        m.currentScreen.setFocus(true)
    end if
end sub

sub ShowSettingsScreen()
    screen = CreateObject("roSGNode", "SettingsScreen")
    if m.userData <> invalid
        screen.userData = m.userData
    end if
    screen.subscriptionActive = (m.subscriptionActive = true)
    screen.observeField("close", "onScreenClose")
    screen.observeField("signedOut", "onSignedOut")
    ShowScreen(screen)
end sub

sub onScreenClose()
    CloseScreen()
end sub

sub onSignedOut()
    ' Clear all screens and go to activation
    while m.screenStack.count() > 0
        old = m.screenStack.pop()
        m.top.findNode("mainSceneId").removeChild(old)
    end while
    if m.currentScreen <> invalid
        m.top.findNode("mainSceneId").removeChild(m.currentScreen)
        m.currentScreen = invalid
    end if
    m.storedToken = ""
    m.userData = invalid
    m.subscriptionActive = false
    ShowActivationScreen()
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press and key = "back"
        if m.screenStack.count() > 0
            CloseScreen()
            return true
        else
            return false
        end if
    end if
    return false
end function
