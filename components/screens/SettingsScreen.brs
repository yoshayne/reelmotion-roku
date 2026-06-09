sub init()
    m.registryTask = CreateObject("roSGNode", "RegistryTask")
    m.registryTask.control = "RUN"

    m.top.observeField("userData", "onUserData")
    m.top.observeField("subscriptionActive", "onSubscriptionActive")

    m.top.findNode("signOutBtn").setFocus(true)

    populateUI()
end sub

sub onUserData()
    populateUI()
end sub

sub onSubscriptionActive()
    populateUI()
end sub

sub populateUI()
    userData = m.top.userData
    subscriptionActive = (m.top.subscriptionActive = true)

    nameLabel = m.top.findNode("nameLabel")
    statusLabel = m.top.findNode("statusLabel")
    joinLabel = m.top.findNode("joinLabel")

    if userData <> invalid
        name = ""
        if userData.name <> invalid and userData.name <> ""
            name = userData.name
        else if userData.email <> invalid
            name = userData.email
        end if
        nameLabel.text = name
    else
        nameLabel.text = "Member"
    end if

    if subscriptionActive
        statusLabel.text = "Active Member"
        statusLabel.color = "#22C55E"
        joinLabel.visible = false
    else
        statusLabel.text = "Not a Member"
        statusLabel.color = "#9CA3AF"
        joinLabel.visible = true
    end if
end sub

sub onSignOut()
    ' Clear registry token
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
    context.observeField("response", "onTokenCleared")
    m.registryTask.request = {context: context}
end sub

sub onTokenCleared()
    m.top.signedOut = true
end sub

sub onShowGuidelines()
    m.top.findNode("guidelinesOverlay").visible = true
    m.top.findNode("guidelinesText").visible = true
    m.top.findNode("closeGuidelinesBtn").visible = true
    m.top.findNode("closeGuidelinesBtn").setFocus(true)
end sub

sub onCloseGuidelines()
    hideGuidelines()
    m.top.findNode("guidelinesBtn").setFocus(true)
end sub

sub hideGuidelines()
    m.top.findNode("guidelinesOverlay").visible = false
    m.top.findNode("guidelinesText").visible = false
    m.top.findNode("closeGuidelinesBtn").visible = false
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            if m.top.findNode("guidelinesOverlay").visible = true
                hideGuidelines()
                m.top.findNode("guidelinesBtn").setFocus(true)
                return true
            else
                m.top.close = true
                return true
            end if
        end if
        if key = "OK"
            if m.top.findNode("guidelinesOverlay").visible = true
                if m.top.findNode("closeGuidelinesBtn").hasFocus()
                    hideGuidelines()
                    m.top.findNode("guidelinesBtn").setFocus(true)
                    return true
                end if
            else
                if m.top.findNode("signOutBtn").hasFocus()
                    onSignOut()
                    return true
                else if m.top.findNode("guidelinesBtn").hasFocus()
                    onShowGuidelines()
                    return true
                end if
            end if
        end if
        if key = "down"
            if m.top.findNode("signOutBtn").hasFocus()
                m.top.findNode("guidelinesBtn").setFocus(true)
                return true
            end if
        end if
        if key = "up"
            if m.top.findNode("guidelinesBtn").hasFocus()
                m.top.findNode("signOutBtn").setFocus(true)
                return true
            end if
        end if
    end if
    return false
end function
