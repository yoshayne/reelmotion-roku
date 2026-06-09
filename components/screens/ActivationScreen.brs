sub init()
    m.registryTask = CreateObject("roSGNode", "RegistryTask")
    m.registryTask.control = "RUN"

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
    if m.pollTask <> invalid
        m.pollTask.active = false
        m.pollTask = invalid
    end if

    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("taskResult", "onPollTaskResult")
    m.pollTask.observeField("status", "onPollStatus")
    m.pollTask.deviceToken = m.deviceToken
    m.pollTask.active = true
    m.pollTask.control = "RUN"
end sub

sub onPollStatus()
    if m.pollTask = invalid then return
    status = m.pollTask.status
    print "ActivationScreen: poll status = " + status

    if status = "pending"
        m.top.findNode("instrLabel").text = "Waiting for activation..." + Chr(10) + "Go to reelmotionapp.com/activate and enter this code"
    else if status = "expired"
        showError("Code expired. Press OK to get a new code.")
    end if
end sub

sub onPollTaskResult()
    if m.pollTask = invalid then return
    result = m.pollTask.taskResult
    print "ActivationScreen: taskResult = " + result

    if result = "activated"
        sessionToken = m.pollTask.sessionToken
        print "ActivationScreen: session token = " + sessionToken
        m.top.findNode("instrLabel").text = "Activating..."
        m.top.findNode("spinner").visible = true
        ' Signal MainScene — it will pick up the token from registry via timer
        m.top.sessionToken = sessionToken
        m.top.activationComplete = true
    else if result = "expired"
        showError("Code expired. Press OK to get a new code.")
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
            if m.pollTask <> invalid
                m.pollTask.active = false
                m.pollTask = invalid
            end if
            requestCode()
            return true
        end if
    end if
    return false
end function
