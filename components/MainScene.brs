sub init()
    m.storedToken = readSessionToken()
    m.userData = invalid
    m.subscriptionActive = false
    m.signInScreen = invalid
    m.homeScreen = invalid
    m.detailScreen = invalid
    m.settingsScreen = invalid
    m.playerScreen = invalid

    m.verifyTask = CreateObject("roSGNode", "HttpTask")
    m.verifyTask.observeField("response", "onVerifyResponse")
    m.verifyTask.control = "RUN"

    m.activationScreen = invalid

    if m.storedToken <> invalid and m.storedToken <> ""
        print "MainScene: session token found, showing home screen"
        showHomeScreen()
        verifyToken(m.storedToken)
    else
        print "MainScene: no session token found, showing sign-in screen"
        showSignInScreen()
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

    m.signInScreen = invalid
    m.activationScreen = invalid
    m.homeScreen = invalid
    m.detailScreen = invalid
    m.settingsScreen = invalid
    m.playerScreen = invalid
end sub

sub showSignInScreen()
    clearScreenStack()

    signIn = CreateObject("roSGNode", "SignInScreen")
    signIn.observeField("signInComplete", "onSignInComplete")
    signIn.observeField("useActivationCode", "onUseActivationCode")
    m.top.appendChild(signIn)
    signIn.visible = true
    signIn.setFocus(true)
    m.signInScreen = signIn

    print "MainScene: showing sign-in screen"
end sub

sub onSignInComplete()
    if m.signInScreen = invalid then return
    if m.signInScreen.signInComplete <> true then return

    token = m.signInScreen.sessionToken
    if token = invalid or token = ""
        token = readSessionToken()
    end if

    if token = invalid or token = ""
        print "MainScene: signInComplete fired without a session token"
        return
    end if

    print "MainScene: sign-in complete, navigating to home"
    m.storedToken = token
    showHomeScreen()
    verifyToken(token)
end sub

sub onUseActivationCode()
    if m.signInScreen = invalid then return
    if m.signInScreen.useActivationCode <> true then return
    print "MainScene: switching to activation screen"
    showActivationScreen()
end sub

sub showActivationScreen()
    clearScreenStack()

    activation = CreateObject("roSGNode", "ActivationScreen")
    activation.observeField("activationComplete", "onActivationComplete")
    activation.observeField("useEmailSignIn", "onUseEmailSignIn")
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
    verifyToken(token)
end sub

sub onUseEmailSignIn()
    if m.activationScreen = invalid then return
    if m.activationScreen.useEmailSignIn <> true then return
    print "MainScene: switching to sign-in screen"
    showSignInScreen()
end sub

sub verifyToken(token as String)
    m.verifyTask.request = {
        url: "https://reelmotionapp.com/api/auth/device/verify",
        method: "GET",
        headers: {Authorization: "Bearer " + token},
        body: "",
        context: "verify"
    }
end sub

sub onVerifyResponse()
    resp = m.verifyTask.response
    if resp = invalid then return
    if resp.context <> "verify" then return

    if resp.code = 200
        json = ParseJson(resp.content)
        if json <> invalid
            if json.subscription_active <> invalid
                m.subscriptionActive = (json.subscription_active = true)
                print "MainScene: subscription_active = " + (m.subscriptionActive).toStr()
                if m.settingsScreen <> invalid
                    m.settingsScreen.subscriptionActive = m.subscriptionActive
                end if
            end if
            if json.user <> invalid
                m.userData = json.user
                if m.settingsScreen <> invalid
                    m.settingsScreen.userData = m.userData
                end if
            end if
        end if
    else if resp.code = 401
        print "MainScene: verify returned 401, clearing token and showing sign-in"
        sec = CreateObject("roRegistrySection", "reelmotion")
        sec.Delete("session_token")
        sec.Flush()
        m.storedToken = ""
        m.userData = invalid
        m.subscriptionActive = false
        showSignInScreen()
    else
        print "MainScene: verify failed with code " + str(resp.code)
    end if
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
    print "MainScene: signed out, clearing registry and returning to sign-in"
    sec = CreateObject("roRegistrySection", "reelmotion")
    sec.Delete("session_token")
    sec.Flush()

    m.storedToken = ""
    m.userData = invalid
    m.subscriptionActive = false
    showSignInScreen()
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
