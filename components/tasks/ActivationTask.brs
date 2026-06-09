sub init()
    m.top.functionName = "runRequest"
    m.top.responseCode = 0
    m.top.responseBody = ""
    m.top.responseJson = invalid
    m.top.errorMessage = ""
    m.top.result = ""
end sub

sub runRequest()
    requestUrl = "https://reelmotionapp.com/api/auth/device/request-code"

    print "ActivationTask: starting request-code POST"
    print "ActivationTask: URL = " + requestUrl

    url = CreateObject("roUrlTransfer")
    url.SetUrl(requestUrl)
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.SetConnectTimeout(30000)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    responseBody = url.GetToString()
    if responseBody = invalid then responseBody = ""

    m.top.responseBody = responseBody
    print "ActivationTask: response body = " + responseBody

    if responseBody <> ""
        json = ParseJson(responseBody)
        if json <> invalid
            m.top.responseCode = 200
            m.top.responseJson = json
            m.top.result = "success"
            return
        end if

        m.top.errorMessage = "Could not parse activation response JSON"
        m.top.result = "error"
        return
    end if

    failureReason = url.GetFailureReason()
    if failureReason = invalid then failureReason = "empty response"
    m.top.errorMessage = failureReason
    print "ActivationTask: request failed = " + failureReason
    m.top.result = "error"
end sub
