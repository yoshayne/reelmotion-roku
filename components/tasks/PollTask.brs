sub init()
    m.top.functionName = "pollLoop"
    m.top.status = "pending"
    m.top.sessionToken = ""
end sub

sub pollLoop()
    while m.top.active = true
        deviceToken = m.top.deviceToken
        if deviceToken = invalid or deviceToken = ""
            sleep(5000)
        else
            doPost(deviceToken)
            if m.top.status = "activated" or m.top.status = "expired"
                return
            end if
            sleep(5000)
        end if
    end while
end sub

sub doPost(deviceToken as String)
    urlXfer = CreateObject("roUrlTransfer")
    urlXfer.SetCertificatesFile("common:/certs/ca-bundle.crt")
    urlXfer.InitClientCertificates()
    urlXfer.SetUrl("https://www.reelmotionapp.com/api/auth/device/poll")
    urlXfer.SetRequest("POST")
    urlXfer.AddHeader("Content-Type", "application/json")

    body = FormatJson({device_token: deviceToken})
    responseStr = urlXfer.PostToString(body)

    if responseStr <> invalid and responseStr <> ""
        json = ParseJson(responseStr)
        if json <> invalid
            status = json.status
            if status = "activated"
                if json.session_token <> invalid
                    m.top.sessionToken = json.session_token
                end if
                m.top.status = "activated"
            else if status = "expired"
                m.top.status = "expired"
            else
                m.top.status = "pending"
            end if
        end if
    end if
end sub
