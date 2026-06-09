sub init()
    m.top.functionName = "runTest"
end sub

sub runTest()
    url = CreateObject("roUrlTransfer")
    url.SetUrl("https://reelmotionapp.com/api/browse-data")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)

    response = url.GetToString()

    if response = invalid or response = ""
        m.top.result = "FAILED: empty response. Reason: " + url.GetFailureReason()
    else if Left(response, 1) = "{"
        m.top.result = "SUCCESS: got JSON response"
    else
        m.top.result = "GOT RESPONSE: " + Left(response, 100)
    end if
end sub
