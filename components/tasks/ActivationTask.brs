sub init()
end sub

sub requestActivationCode()
    url = CreateObject("roUrlTransfer")
    url.SetUrl("https://reelmotionapp.com/api/auth/device/request-code")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    response = url.PostFromString("{}")
    responseBody = url.GetToString()

    print "ActivationTask: response code = " + str(response)
    print "ActivationTask: response body = " + responseBody

    if responseBody = invalid or responseBody = ""
        m.top.errorMessage = "No response from server. Reason: " + url.GetFailureReason()
        return
    end if

    json = ParseJson(responseBody)
    if json = invalid
        m.top.errorMessage = "Invalid response: " + Left(responseBody, 100)
        return
    end if

    if json.code = invalid
        m.top.errorMessage = "No code in response: " + Left(responseBody, 100)
        return
    end if

    m.top.deviceToken = json.device_token
    m.top.code = json.code
end sub
