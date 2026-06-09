sub init()
    m.port = CreateObject("roMessagePort")
    m.top.observeField("request", m.port)
    m.top.functionName = "go"
    m.top.control = "RUN"
end sub

sub go()
    while true
        msg = wait(0, m.port)
        if type(msg) = "roSGNodeEvent"
            if msg.getField() = "request"
                doRequest(m.top.request)
            end if
        end if
    end while
end sub

sub doRequest(req as Object)
    if req = invalid then return

    url = req.url
    method = req.method
    headers = req.headers
    body = req.body
    context = req.context

    if url = invalid or url = "" then
        print "HttpTask: ERROR - empty URL"
        m.top.response = {code: -1, content: "", context: context}
        return
    end if

    if method = invalid then method = "GET"

    print "HttpTask: >>> " + UCase(method) + " " + url

    urlXfer = CreateObject("roUrlTransfer")
    urlXfer.SetCertificatesFile("common:/certs/ca-bundle.crt")
    urlXfer.InitClientCertificates()
    urlXfer.EnableEncodings(true)
    urlXfer.RetainBodyOnError(true)
    urlXfer.SetConnectTimeout(60000)
    urlXfer.SetUrl(url)
    urlXfer.SetPort(m.port)

    if UCase(method) = "POST"
        urlXfer.SetRequest("POST")
        urlXfer.AddHeader("Content-Type", "application/json")
        if headers <> invalid
            for each key in headers
                urlXfer.AddHeader(key, headers[key])
            end for
        end if
        postBody = ""
        if body <> invalid then postBody = body
        print "HttpTask: POST body = " + postBody
        ok = urlXfer.AsyncPostFromString(postBody)
    else
        urlXfer.SetRequest("GET")
        if headers <> invalid
            for each key in headers
                urlXfer.AddHeader(key, headers[key])
            end for
        end if
        ok = urlXfer.AsyncGetToString()
    end if

    if not ok
        reason = urlXfer.GetFailureReason()
        print "HttpTask: ERROR - AsyncRequest failed, reason: " + reason
        m.top.response = {code: -1, content: "", context: context}
        return
    end if

    responseMsg = wait(60000, m.port)
    if type(responseMsg) = "roUrlEvent"
        code = responseMsg.GetResponseCode()
        content = responseMsg.GetString()
        reason = urlXfer.GetFailureReason()
        print "HttpTask: <<< response code = " + str(code)
        print "HttpTask: <<< response body = " + content
        if reason <> "" and reason <> invalid
            print "HttpTask: <<< failure reason = " + reason
        end if
        m.top.response = {code: code, content: content, context: context}
    else
        reason = urlXfer.GetFailureReason()
        print "HttpTask: ERROR - wait timed out or wrong event type: " + type(responseMsg)
        print "HttpTask: <<< failure reason = " + reason
        m.top.response = {code: -1, content: "", context: context}
    end if
end sub
