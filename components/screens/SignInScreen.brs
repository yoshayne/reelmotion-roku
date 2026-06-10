sub init()
    m.focusedField = "email"
    m.busy = false

    m.signInTask = CreateObject("roSGNode", "HttpTask")
    m.signInTask.observeField("response", "onSignInResponse")
    m.signInTask.control = "RUN"

    m.activationTask = invalid
    m.pollTask = invalid

    emailField = m.top.findNode("emailField")
    if emailField <> invalid then emailField.setFocus(true)
    updateFocusBorders()

    requestActivationCode()
end sub

sub updateFocusBorders()
    m.top.findNode("emailFocusBorder").visible = (m.focusedField = "email")
    m.top.findNode("passwordFocusBorder").visible = (m.focusedField = "password")
    m.top.findNode("signInBtnFocusBorder").visible = (m.focusedField = "signin")
    m.top.findNode("retryBtnFocusBorder").visible = (m.focusedField = "retry")
end sub

sub showError(msg as String)
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("spinner").visible = false
    m.busy = false
end sub

sub clearError()
    m.top.findNode("errorLabel").visible = false
end sub

sub doSignIn()
    emailField = m.top.findNode("emailField")
    passwordField = m.top.findNode("passwordField")

    email = emailField.text
    password = passwordField.text

    if email = invalid then email = ""
    if password = invalid then password = ""

    email = email.trim()

    if email = ""
        showError("Please enter your email address.")
        m.top.findNode("emailField").setFocus(true)
        m.focusedField = "email"
        updateFocusBorders()
        return
    end if

    if password = ""
        showError("Please enter your password.")
        m.top.findNode("passwordField").setFocus(true)
        m.focusedField = "password"
        updateFocusBorders()
        return
    end if

    clearError()
    m.busy = true
    m.top.findNode("spinner").visible = true

    body = FormatJson({email: email, password: password})

    m.signInTask.request = {
        url: "https://reelmotionapp.com/api/auth/clerk-session",
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Accept": "application/json"
        },
        body: body,
        context: "signIn"
    }
end sub

sub onSignInResponse()
    resp = m.signInTask.response
    if resp = invalid then return
    if resp.context <> "signIn" then return

    m.top.findNode("spinner").visible = false
    m.busy = false

    if resp.code = 200
        json = ParseJson(resp.content)
        if json <> invalid and json.sessionToken <> invalid and json.sessionToken <> ""
            token = json.sessionToken
            sec = CreateObject("roRegistrySection", "reelmotion")
            sec.Write("session_token", token)
            sec.Flush()
            print "SignInScreen: session token saved to registry"
            m.top.sessionToken = token
            m.top.signInComplete = true
        else
            showError("Sign in failed. Please try again.")
        end if
    else if resp.code = 401
        json = ParseJson(resp.content)
        errMsg = "Invalid email or password."
        if json <> invalid and json.error <> invalid and json.error <> ""
            errMsg = json.error
        end if
        showError(errMsg)
    else if resp.code = -1
        showError("Could not connect. Check your network and try again.")
    else
        showError("Sign in failed (error " + str(resp.code).trim() + "). Please try again.")
    end if
end sub

' ===================== Activation (right column) =====================

sub requestActivationCode()
    m.top.findNode("activationErrorLabel").visible = false
    m.top.findNode("retryBtnBg").visible = false
    m.top.findNode("retryBtnFocusBorder").visible = false
    m.top.findNode("retryBtnLabel").visible = false
    m.top.findNode("activationSpinner").visible = true
    m.top.findNode("codeLabel").text = ""

    if m.pollTask <> invalid
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if

    m.activationTask = CreateObject("roSGNode", "ActivationTask")
    m.activationTask.observeField("code", "onCodeReceived")
    m.activationTask.observeField("errorMessage", "onActivationError")
    m.activationTask.functionName = "requestActivationCode"
    m.activationTask.control = "RUN"
end sub

sub onCodeReceived()
    if m.activationTask = invalid then return
    code = m.activationTask.code
    if code = invalid or code = "" then return

    m.top.findNode("codeLabel").text = code
    m.top.findNode("activationSpinner").visible = false

    deviceToken = m.activationTask.deviceToken
    if deviceToken <> invalid and deviceToken <> ""
        startPolling(deviceToken)
    end if
end sub

sub onActivationError()
    if m.activationTask = invalid then return
    errMsg = m.activationTask.errorMessage
    if errMsg = invalid or errMsg = "" then return
    showActivationError(errMsg)
end sub

sub startPolling(deviceToken as String)
    if m.pollTask <> invalid
        m.pollTask.control = "STOP"
        m.pollTask = invalid
    end if

    m.pollTask = CreateObject("roSGNode", "PollTask")
    m.pollTask.observeField("sessionToken", "onActivationSessionToken")
    m.pollTask.observeField("codeExpired", "onCodeExpired")
    m.pollTask.deviceToken = deviceToken
    m.pollTask.functionName = "pollForActivation"
    m.pollTask.control = "RUN"
end sub

sub onActivationSessionToken()
    if m.pollTask = invalid then return
    token = m.pollTask.sessionToken
    if token = invalid or token = "" then return

    sec = CreateObject("roRegistrySection", "reelmotion")
    sec.Write("session_token", token)
    sec.Flush()
    print "SignInScreen: session token saved to registry (activation)"

    m.top.sessionToken = token
    m.top.signInComplete = true
end sub

sub onCodeExpired()
    if m.pollTask = invalid then return
    if m.pollTask.codeExpired <> true then return
    showActivationError("Code expired. Press OK to get a new code.")
end sub

sub showActivationError(msg as String)
    m.top.findNode("activationSpinner").visible = false
    m.top.findNode("codeLabel").text = ""
    m.top.findNode("activationErrorLabel").text = msg
    m.top.findNode("activationErrorLabel").visible = true
    m.top.findNode("retryBtnBg").visible = true
    m.top.findNode("retryBtnLabel").visible = true
end sub

' ===================== Key handling =====================

function onKeyEvent(key as String, press as Boolean) as Boolean
    if not press then return false

    if key = "down"
        if m.focusedField = "email"
            m.top.findNode("passwordField").setFocus(true)
            m.focusedField = "password"
            updateFocusBorders()
            return true
        else if m.focusedField = "password"
            m.focusedField = "signin"
            m.top.setFocus(true)
            updateFocusBorders()
            return true
        end if
    end if

    if key = "up"
        if m.focusedField = "signin"
            m.top.findNode("passwordField").setFocus(true)
            m.focusedField = "password"
            updateFocusBorders()
            return true
        else if m.focusedField = "password"
            m.top.findNode("emailField").setFocus(true)
            m.focusedField = "email"
            updateFocusBorders()
            return true
        end if
    end if

    if key = "right"
        if m.focusedField = "signin" and m.top.findNode("retryBtnLabel").visible = true
            m.focusedField = "retry"
            m.top.setFocus(true)
            updateFocusBorders()
            return true
        end if
    end if

    if key = "left"
        if m.focusedField = "retry"
            m.focusedField = "signin"
            m.top.setFocus(true)
            updateFocusBorders()
            return true
        end if
    end if

    if key = "OK"
        if m.busy then return true
        if m.focusedField = "signin"
            doSignIn()
            return true
        else if m.focusedField = "retry"
            requestActivationCode()
            return true
        end if
        ' email/password: let TextEditBox handle OK to open the on-screen keyboard
        return false
    end if

    return false
end function
