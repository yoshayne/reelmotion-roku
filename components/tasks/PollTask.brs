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

        body = FormatJson({ deviceToken: deviceToken })
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
