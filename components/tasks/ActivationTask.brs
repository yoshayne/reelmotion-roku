sub init()
end sub

sub requestActivationCode()
    port = CreateObject("roMessagePort")
    url = CreateObject("roUrlTransfer")
    url.SetMessagePort(port)
    url.SetUrl("https://reelmotionapp.com/api/auth/device/request-code")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    print "ActivationTask: sending request-code POST"

    if not url.AsyncPostFromString("{}")
        m.top.errorMessage = "Request failed to start: " + url.GetFailureReason()
        return
    end if

    msg = wait(30000, port)
    if type(msg) <> "roUrlEvent"
        m.top.errorMessage = "Request timed out or invalid event: " + type(msg)
        return
    end if

    responseCode = msg.GetResponseCode()
    responseBody = msg.GetString()
    failureReason = msg.GetFailureReason()

    print "ActivationTask: response code = " + str(responseCode)
    print "ActivationTask: response body = " + responseBody

    if failureReason <> "" and failureReason <> invalid
        print "ActivationTask: failure reason = " + failureReason
    end if

    if responseBody = invalid or responseBody = ""
        m.top.errorMessage = "Empty response (code " + str(responseCode).trim() + ")"
        return
    end if

    json = ParseJson(responseBody)
    if json = invalid
        m.top.errorMessage = "Invalid JSON: " + Left(responseBody, 100)
        return
    end if

    if json.code = invalid
        m.top.errorMessage = "No code field in response: " + Left(responseBody, 100)
        return
    end if

    deviceToken = json.device_token
    if deviceToken = invalid then deviceToken = ""

    m.top.deviceToken = deviceToken
    m.top.code = json.code
end sub
