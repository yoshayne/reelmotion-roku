sub init()
    m.top.functionName = "pollLoop"
    m.top.status = "pending"
    m.top.taskResult = ""
    m.top.sessionToken = ""
end sub

sub pollLoop()
    while m.top.active = true
        deviceToken = m.top.deviceToken
        if deviceToken = invalid or deviceToken = ""
            sleep(5000)
        else
            doPost(deviceToken)
            if m.top.taskResult = "activated" or m.top.status = "expired"
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
    urlXfer.EnableEncodings(true)
    urlXfer.RetainBodyOnError(true)
    urlXfer.SetConnectTimeout(30000)
    urlXfer.SetUrl("https://reelmotionapp.com/api/auth/device/poll")
    urlXfer.SetRequest("POST")
    urlXfer.AddHeader("Content-Type", "application/json")
    urlXfer.AddHeader("Accept", "application/json")

    body = FormatJson({device_token: deviceToken})
    responseCode = urlXfer.PostFromString(body)
    responseStr = urlXfer.GetToString()

    print "PollTask: poll response code = " + str(responseCode)
    print "PollTask: poll response body = " + responseStr

    if responseCode = 200 and responseStr <> invalid and responseStr <> ""
        json = ParseJson(responseStr)
        if json <> invalid
            status = json.status
            print "PollTask: status = " + status
            if status = "activated"
                token = json.session_token
                if token = invalid then token = ""
                m.top.sessionToken = token
                m.top.status = "activated"
                ' Write token to registry directly from task thread
                writeTokenToRegistry(token)
                ' Fire taskResult last — this is what the timer and observers watch
                m.top.taskResult = "activated"
            else if status = "expired"
                m.top.status = "expired"
                m.top.taskResult = "expired"
            else
                m.top.status = "pending"
            end if
        end if
    else
        print "PollTask: non-200 or empty response, will retry"
    end if
end sub

sub writeTokenToRegistry(token as String)
    if token = invalid or token = "" then return
    print "PollTask: writing session_token to registry section reelmotion"
    sec = CreateObject("roRegistrySection", "reelmotion")
    sec.Write("session_token", token)
    sec.Flush()
    print "PollTask: registry write complete"
end sub
