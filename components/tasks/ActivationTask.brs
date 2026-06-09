sub init()
    m.top.functionName = "runRequest"
    m.top.responseCode = 0
    m.top.done = false
end sub

sub runRequest()
    json = RequestActivationCode()
    if json <> invalid
        m.top.responseJson = json
    end if
    m.top.done = true
end sub

function RequestActivationCode() as Object
    url = CreateObject("roUrlTransfer")
    url.SetUrl("https://reelmotionapp.com/api/auth/device/request-code")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    responseCode = url.PostFromString("{}")
    responseBody = url.GetToString()

    print "Activation request response code: " + str(responseCode)
    print "Activation request response body: " + responseBody

    m.top.responseCode = responseCode

    if responseCode = 200
        return ParseJson(responseBody)
    end if
    return invalid
end function
