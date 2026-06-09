sub init()
end sub

sub pollForActivation()
    deviceToken = m.top.deviceToken
    if deviceToken = invalid or deviceToken = ""
        print "PollTask: no device token, aborting"
        return
    end if

    print "PollTask: starting poll loop for token " + deviceToken

    while true
        Sleep(5000)

        port = CreateObject("roMessagePort")
        url = CreateObject("roUrlTransfer")
        url.SetMessagePort(port)
        url.SetUrl("https://reelmotionapp.com/api/auth/device/poll")
        url.SetCertificatesFile("common:/certs/ca-bundle.crt")
        url.InitClientCertificates()
        url.EnableEncodings(true)
        url.RetainBodyOnError(true)
        url.AddHeader("Content-Type", "application/json")
        url.AddHeader("Accept", "application/json")
        url.SetRequest("POST")

        body = FormatJson({device_token: deviceToken})

        if not url.AsyncPostFromString(body)
            print "PollTask: request failed to start: " + url.GetFailureReason()
        else
            msg = wait(30000, port)
            if type(msg) = "roUrlEvent"
                responseCode = msg.GetResponseCode()
                responseBody = msg.GetString()

                print "PollTask: response code = " + str(responseCode)
                print "PollTask: response body = " + responseBody

                if responseBody <> invalid and responseBody <> ""
                    json = ParseJson(responseBody)
                    if json <> invalid and json.status <> invalid
                        if json.status = "activated"
                            sessionToken = json.session_token
                            if sessionToken = invalid then sessionToken = ""
                            m.top.sessionToken = sessionToken
                            return
                        else if json.status = "expired"
                            m.top.codeExpired = true
                            return
                        end if
                    end if
                end if
            else
                print "PollTask: wait timed out or wrong event type: " + type(msg)
            end if
        end if
    end while
end sub
