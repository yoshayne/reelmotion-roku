sub init()
    m.focusedField = "email"
    m.busy = false

    m.signInTask = CreateObject("roSGNode", "HttpTask")
    m.signInTask.observeField("response", "onSignInResponse")
    m.signInTask.functionName = "go"
    m.signInTask.control = "RUN"

    emailField = m.top.findNode("emailField")
    if emailField <> invalid then emailField.setFocus(true)
    updateFocusBorders()
end sub

sub updateFocusBorders()
    m.top.findNode("emailFocusBorder").visible = (m.focusedField = "email")
    m.top.findNode("passwordFocusBorder").visible = (m.focusedField = "password")
end sub

sub showError(msg as String)
    m.top.findNode("errorLabel").text = msg
    m.top.findNode("errorLabel").visible = true
    m.top.findNode("spinner").visible = false
    m.top.findNode("signInBtnLabel").visible = true
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
    m.top.findNode("signInBtnLabel").visible = false

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
    m.top.findNode("signInBtnLabel").visible = true
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

function onKeyEvent(key as String, press as Boolean) as Boolean
    if not press then return false

    if key = "down"
        if m.focusedField = "email"
            m.top.findNode("passwordField").setFocus(true)
            m.focusedField = "password"
            updateFocusBorders()
            return true
        else if m.focusedField = "password"
            m.focusedField = "button"
            updateFocusBorders()
            return true
        else if m.focusedField = "button"
            m.top.useActivationCode = true
            return true
        end if
    end if

    if key = "up"
        if m.focusedField = "button"
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

    if key = "OK"
        if m.busy then return true
        if m.focusedField = "email"
            m.top.findNode("passwordField").setFocus(true)
            m.focusedField = "password"
            updateFocusBorders()
            return true
        else
            doSignIn()
            return true
        end if
    end if

    return false
end function
