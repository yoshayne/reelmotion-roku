sub init()
    m.activationTask = invalid
    m.pollTask = invalid

    m.activationTask = CreateObject("roSGNode", "ActivationTask")
    m.activationTask.observeField("code", "onCodeReceived")
    m.activationTask.observeField("errorMessage", "onActivationError")
    m.activationTask.functionName = "requestActivationCode"
    m.activationTask.control = "RUN"
end sub

sub onCodeReceived()
    code = m.activationTask.code
    if code = invalid or code = "" then return

    m.top.findNode("codeLabel").text = code
    m.top.findNode("instrLabel").text = "Go to reelmotionapp.com/activate" + Chr(10) + "and enter this code"
    m.top.findNode("spinner").visible = false

    deviceToken = m.activationTask.deviceToken
    if deviceToken <> invalid and deviceToken <> ""
        startPolling(deviceToken)
    end if
end sub

sub onActivationError()
    errMsg = m.activationTask.errorMessage
    if errMsg = invalid then errMsg = "Could not connect. Please retry."
    showError(errMsg)
end sub

sub startPolling(deviceToken as String)
    if m.pollTask <> invalid
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if

    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("sessionToken", "onSessionTokenReceived")
    m.pollTask.observeField("codeExpired", "onCodeExpired")
    m.pollTask.deviceToken = deviceToken
    m.pollTask.functionName = "pollForActivation"
    m.pollTask.control = "RUN"
end sub

sub onSessionTokenReceived()
    token = m.pollTask.sessionToken
    if token = invalid or token = "" then return

    sec = CreateObject("roRegistrySection", "reelmotion")
    sec.Write("session_token", token)
    sec.Flush()
    print "ActivationScreen: session token saved to registry"

    m.top.findNode("instrLabel").text = "Activation complete. Loading..."
    m.top.findNode("spinner").visible = true

    m.top.sessionToken = token
    m.top.activationComplete = true
end sub

sub onCodeExpired()
    if m.pollTask = invalid then return
    if m.pollTask.codeExpired <> true then return
    showError("Code expired. Press OK to get a new code.")
end sub

sub showError(msg as String)
    m.top.findNode("spinner").visible = false
    m.top.findNode("codeLabel").text = ""
    m.top.findNode("instrLabel").text = ""
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("retryButton").visible = true
    m.top.findNode("retryButton").setFocus(true)
end sub

sub retryActivation()
    if m.pollTask <> invalid
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if
    m.activationTask = invalid

    m.top.findNode("codeLabel").text = ""
    m.top.findNode("instrLabel").text = "Requesting activation code..."
    m.top.findNode("errorLabel").visible = false
    m.top.findNode("retryButton").visible = false
    m.top.findNode("spinner").visible = true

    m.activationTask = CreateObject("roSGNode", "ActivationTask")
    m.activationTask.observeField("code", "onCodeReceived")
    m.activationTask.observeField("errorMessage", "onActivationError")
    m.activationTask.functionName = "requestActivationCode"
    m.activationTask.control = "RUN"
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if not press then return false

    if key = "OK"
        retryBtn = m.top.findNode("retryButton")
        if retryBtn.visible = true
            retryActivation()
            return true
        end if
    end if

    if key = "down"
        m.top.useEmailSignIn = true
        return true
    end if

    return false
end function
