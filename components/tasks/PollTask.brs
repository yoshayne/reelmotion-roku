sub init()
end sub

sub pollForActivation()
    deviceToken = m.top.deviceToken

    while true
        Sleep(5000)

        url = CreateObject("roUrlTransfer")
        url.SetUrl("https://reelmotionapp.com/api/auth/device/poll")
        url.SetCertificatesFile("common:/certs/ca-bundle.crt")
        url.InitClientCertificates()
        url.EnableEncodings(true)
        url.RetainBodyOnError(true)
        url.AddHeader("Content-Type", "application/json")
        url.AddHeader("Accept", "application/json")
        url.SetRequest("POST")

        body = FormatJson({device_token: deviceToken})
        response = url.PostFromString(body)
        responseBody = url.GetToString()

        print "PollTask: response code = " + str(response)
        print "PollTask: response body = " + responseBody

        if responseBody <> invalid and responseBody <> ""
            json = ParseJson(responseBody)
            if json <> invalid
                if json.status = "activated"
                    m.top.sessionToken = json.session_token
                    return
                else if json.status = "expired"
                    m.top.codeExpired = true
                    return
                end if
            end if
        end if
    end while
end sub
