sub init()
    m.activationTask = invalid
    m.deviceToken = ""
    m.pollTask = invalid
    m.debugLines = []

    appendDebugLog("ActivationScreen: initialized")
    requestCode()
end sub

sub requestCode()
    appendDebugLog("ActivationScreen: requesting activation code")

    m.top.findNode("codeLabel").text = ""
    m.top.findNode("errorLabel").visible = false
    m.top.findNode("retryButton").visible = false
    m.top.findNode("spinner").visible = true
    m.top.findNode("instrLabel").text = "Requesting activation code..."

    if m.activationTask <> invalid
        m.activationTask = invalid
    end if

    m.activationTask = CreateObject("roSGNode", "ActivationTask")
    m.activationTask.observeField("done", "onActivationDone")
    m.activationTask.control = "RUN"
end sub

sub onActivationDone()
    if m.activationTask = invalid then return

    code = m.activationTask.responseCode
    json = m.activationTask.responseJson

    print "ActivationScreen: task done, HTTP code = " + str(code)
    appendDebugLog("ActivationScreen: activation code HTTP " + str(code).trim())

    if code = 200 and json <> invalid and json.code <> invalid
        m.deviceToken = json.device_token
        if m.deviceToken = invalid then m.deviceToken = ""
        appendDebugLog("ActivationScreen: received code and device token")
        m.top.findNode("codeLabel").text = json.code
        m.top.findNode("instrLabel").text = "Waiting for activation..." + Chr(10) + "Go to reelmotionapp.com/activate and enter this code"
        m.top.findNode("spinner").visible = true
        startPolling()
    else
        appendDebugLog("ActivationScreen: failed to get activation code")
        showError("Error: HTTP " + str(code).trim() + " — could not get activation code. Press OK to retry.")
    end if
end sub

sub startPolling()
    stopPolling()

    appendDebugLog("ActivationScreen: starting PollTask")
    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("activationComplete", "onPollActivationComplete")
    m.pollTask.observeField("codeExpired", "onPollCodeExpired")
    m.pollTask.observeField("logMessage", "onPollLogMessage")
    m.pollTask.deviceToken = m.deviceToken
    m.pollTask.control = "RUN"
end sub

sub onPollActivationComplete()
    if m.pollTask = invalid then return
    if m.pollTask.activationComplete <> true then return

    sessionToken = m.pollTask.sessionToken
    if sessionToken = invalid then sessionToken = ""

    print "ActivationScreen: PollTask activation complete"
    appendDebugLog("ActivationScreen: PollTask activation complete")
    m.top.findNode("instrLabel").text = "Activation complete. Loading..."
    m.top.findNode("spinner").visible = true

    m.top.sessionToken = sessionToken
    m.top.activationComplete = true
end sub

sub onPollCodeExpired()
    if m.pollTask = invalid then return
    if m.pollTask.codeExpired = true
        appendDebugLog("ActivationScreen: activation code expired")
        showError("Code expired. Press OK to get a new code.")
    end if
end sub

sub onPollLogMessage()
    if m.pollTask = invalid then return
    appendDebugLog(m.pollTask.logMessage)
end sub

sub stopPolling()
    if m.pollTask <> invalid
        appendDebugLog("ActivationScreen: stopping PollTask")
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if
end sub

sub showError(msg as String)
    appendDebugLog("ActivationScreen ERROR: " + msg)

    m.top.findNode("spinner").visible = false
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("retryButton").visible = true
    m.top.findNode("retryButton").setFocus(true)
end sub

sub appendDebugLog(message as String)
    if message = invalid or message = "" then return

    print message

    if m.debugLines = invalid
        m.debugLines = []
    end if

    m.debugLines.push(message)
    while m.debugLines.count() > 6
        m.debugLines.shift()
    end while

    logText = "Debug log:"
    for each line in m.debugLines
        logText = logText + Chr(10) + line
    end for

    debugLabel = m.top.findNode("debugLabel")
    if debugLabel <> invalid
        debugLabel.text = logText
    end if
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press and key = "OK"
        retryBtn = m.top.findNode("retryButton")
        if retryBtn <> invalid and retryBtn.visible = true
            stopPolling()
            requestCode()
            return true
        end if
    end if
    return false
end function
