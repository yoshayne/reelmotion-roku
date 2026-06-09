sub init()
    m.activationTask = invalid
    m.deviceToken = ""
    m.pollTask = invalid

    requestCode()
end sub

sub requestCode()
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

    if code = 200 and json <> invalid and json.code <> invalid
        m.deviceToken = json.device_token
        m.top.findNode("codeLabel").text = json.code
        m.top.findNode("instrLabel").text = "Waiting for activation..." + Chr(10) + "Go to reelmotionapp.com/activate and enter this code"
        m.top.findNode("spinner").visible = true
        startPolling()
    else
        showError("Error: HTTP " + str(code).trim() + " — could not get activation code. Press OK to retry.")
    end if
end sub

sub startPolling()
    stopPolling()

    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("activationComplete", "onPollActivationComplete")
    m.pollTask.observeField("codeExpired", "onPollCodeExpired")
    m.pollTask.deviceToken = m.deviceToken
    m.pollTask.control = "RUN"
end sub

sub onPollActivationComplete()
    if m.pollTask = invalid then return
    if m.pollTask.activationComplete <> true then return

    sessionToken = m.pollTask.sessionToken
    if sessionToken = invalid then sessionToken = ""

    print "ActivationScreen: PollTask activation complete"
    m.top.findNode("instrLabel").text = "Activation complete. Loading..."
    m.top.findNode("spinner").visible = true

    m.top.sessionToken = sessionToken
    m.top.activationComplete = true
end sub

sub onPollCodeExpired()
    if m.pollTask = invalid then return
    if m.pollTask.codeExpired = true
        showError("Code expired. Press OK to get a new code.")
    end if
end sub

sub stopPolling()
    if m.pollTask <> invalid
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if
end sub

sub showError(msg as String)
    m.top.findNode("spinner").visible = false
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("retryButton").visible = true
    m.top.findNode("retryButton").setFocus(true)
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
