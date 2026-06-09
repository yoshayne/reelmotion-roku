sub init()
    m.top.functionName = "runPoll"
    m.top.sessionToken = ""
    m.top.activationComplete = false
    m.top.codeExpired = false
end sub

sub runPoll()
    deviceToken = m.top.deviceToken
    baseUrl = "https://reelmotionapp.com"

    if deviceToken = invalid or deviceToken = ""
        print "PollTask: missing device token; cannot start polling"
        return
    end if

    print "PollTask: started polling with device token"

    while true
        ' Wait 5 seconds between polls.
        sleep(5000)

        response = postPollRequest(baseUrl, deviceToken)
        responseCode = response.code
        responseBody = response.body

        print "PollTask: poll response code = " + str(responseCode).trim()
        print "PollTask: poll response body = " + responseBody

        if responseBody <> ""
            json = ParseJson(responseBody)
            if json <> invalid and json.status <> invalid
                if json.status = "activated"
                    sessionToken = json.session_token
                    if sessionToken = invalid then sessionToken = ""

                    if sessionToken <> ""
                        sec = CreateObject("roRegistrySection", "reelmotion")
                        sec.Write("session_token", sessionToken)
                        sec.Flush()
                    end if

                    m.top.sessionToken = sessionToken
                    m.top.activationComplete = true
                    return
                else if json.status = "expired"
                    m.top.codeExpired = true
                    return
                else
                    print "PollTask: activation still pending"
                end if
            else
                print "PollTask: could not parse poll response JSON"
            end if
        else
            print "PollTask: empty response; polling will retry"
        end if
    end while
end sub

function postPollRequest(baseUrl as String, deviceToken as String) as Object
    port = CreateObject("roMessagePort")
    url = CreateObject("roUrlTransfer")
    url.SetMessagePort(port)
    url.SetUrl(baseUrl + "/api/auth/device/poll")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.SetConnectTimeout(30000)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    body = FormatJson({ device_token: deviceToken })
    print "PollTask: posting to " + baseUrl + "/api/auth/device/poll"

    if url.AsyncPostFromString(body) <> true
        print "PollTask: AsyncPostFromString failed to start: " + url.GetFailureReason()
        return { code: -1, body: "" }
    end if

    event = wait(30000, port)
    if event = invalid
        url.AsyncCancel()
        print "PollTask: poll request timed out"
        return { code: -1, body: "" }
    end if

    if type(event) <> "roUrlEvent"
        print "PollTask: unexpected poll event type = " + type(event)
        return { code: -1, body: "" }
    end if

    responseBody = event.GetString()
    if responseBody = invalid then responseBody = ""

    failureReason = event.GetFailureReason()
    if failureReason <> invalid and failureReason <> ""
        print "PollTask: poll failure reason = " + failureReason
    end if

    return {
        code: event.GetResponseCode(),
        body: responseBody
    }
end function
    while true
        ' Wait 5 seconds between polls.
        sleep(5000)

        url = CreateObject("roUrlTransfer")
        url.SetUrl(baseUrl + "/api/auth/device/poll")
        url.SetCertificatesFile("common:/certs/ca-bundle.crt")
        url.InitClientCertificates()
        url.EnableEncodings(true)
        url.RetainBodyOnError(true)
        url.SetConnectTimeout(30000)
        url.AddHeader("Content-Type", "application/json")
        url.AddHeader("Accept", "application/json")
        url.SetRequest("POST")

        body = FormatJson({ device_token: deviceToken })
        responseCode = url.PostFromString(body)
        responseBody = url.GetToString()
        if responseBody = invalid then responseBody = ""

        print "PollTask: poll response code = " + str(responseCode).trim()
        print "PollTask: poll response body = " + responseBody

        if responseBody <> invalid and responseBody <> ""
            json = ParseJson(responseBody)
            if json <> invalid and json.status <> invalid
                if json.status = "activated"
                    sessionToken = json.session_token
                    if sessionToken = invalid then sessionToken = ""

                    if sessionToken <> ""
                        sec = CreateObject("roRegistrySection", "reelmotion")
                        sec.Write("session_token", sessionToken)
                        sec.Flush()
                    end if

                    m.top.sessionToken = sessionToken
                    m.top.activationComplete = true
                    return
                else if json.status = "expired"
                    m.top.codeExpired = true
                    return
                end if
            end if
        else
            print "PollTask: empty response; polling will retry"
        end if
    end while
end sub
