sub init()
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onHttpResponse")

    m.registryTask = CreateObject("roSGNode", "RegistryTask")
    m.registryTask.control = "RUN"

    m.pendingRequest = ""
    m.deviceToken = ""
    m.pollTask = invalid

    connectivityTest()
    requestCode()
end sub

sub connectivityTest()
    print "ActivationScreen: running connectivity test to https://reelmotionapp.com"
    m.httpTask.request = {
        url: "https://reelmotionapp.com",
        method: "GET",
        headers: {},
        body: "",
        context: "connectivityTest"
    }
end sub

sub requestCode()
    ' Reset UI
    m.top.findNode("codeLabel").text = ""
    m.top.findNode("errorLabel").visible = false
    m.top.findNode("retryButton").visible = false
    m.top.findNode("spinner").visible = true
    m.top.findNode("instrLabel").text = "Requesting activation code..."

    m.pendingRequest = "requestCode"
    m.httpTask.request = {
        url: "https://reelmotionapp.com/api/auth/device/request-code",
        method: "POST",
        headers: {"Content-Type": "application/json"},
        body: "{}",
        context: "requestCode"
    }
end sub

sub onHttpResponse()
    resp = m.httpTask.response
    if resp = invalid then
        showError("Network error. Please try again.")
        return
    end if

    context = resp.context

    if context = "connectivityTest"
        print "ActivationScreen: connectivity test response code = " + str(resp.code)
        print "ActivationScreen: connectivity test body = " + resp.content
        requestCode()
        return
    end if

    if context = "requestCode"
        if resp.code = 200
            json = ParseJson(resp.content)
            if json <> invalid and json.code <> invalid
                m.deviceToken = json.device_token
                m.top.findNode("codeLabel").text = json.code
                m.top.findNode("instrLabel").text = "Go to reelmotionapp.com/activate on your phone or computer and enter this code"
                m.top.findNode("spinner").visible = false
                startPolling()
            else
                showError("Invalid response from server. Please try again.")
            end if
        else
            showError("Could not get activation code (Error " + str(resp.code).trim() + "). Please try again.")
        end if
    end if
end sub

sub startPolling()
    if m.pollTask <> invalid
        m.pollTask.active = false
        m.pollTask = invalid
    end if

    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("status", "onPollStatus")
    m.pollTask.deviceToken = m.deviceToken
    m.pollTask.active = true
    m.pollTask.control = "RUN"
end sub

sub onPollStatus()
    if m.pollTask = invalid then return
    status = m.pollTask.status

    if status = "activated"
        sessionToken = m.pollTask.sessionToken
        if sessionToken <> invalid and sessionToken <> ""
            saveToken(sessionToken)
        end if
    else if status = "expired"
        showError("Code expired. Please get a new code.")
    end if
end sub

sub saveToken(token as String)
    context = CreateObject("roSGNode", "Node")
    context.addFields({
        parameters: {
            command: "write",
            section: "reelmotion",
            key: "device_token",
            value: token
        },
        response: {}
    })
    context.observeField("response", "onTokenSaved")
    m.registryTask.request = {context: context}
end sub

sub onTokenSaved()
    m.top.goHome = true
end sub

sub showError(msg as String)
    m.top.findNode("spinner").visible = false
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("retryButton").visible = true
    m.top.findNode("retryButton").setFocus(true)
end sub

sub onRetrySelected()
    if m.pollTask <> invalid
        m.pollTask.active = false
        m.pollTask = invalid
    end if
    requestCode()
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press and key = "OK"
        retryBtn = m.top.findNode("retryButton")
        if retryBtn <> invalid and retryBtn.visible = true
            onRetrySelected()
            return true
        end if
    end if
    return false
end function
