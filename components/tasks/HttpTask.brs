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
        m.top.response = {code: -1, content: "", context: context}
        return
    end if

    if method = invalid then method = "GET"

    urlXfer = CreateObject("roUrlTransfer")
    urlXfer.SetCertificatesFile("common:/certs/ca-bundle.crt")
    urlXfer.InitClientCertificates()
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

    if ok
        responseMsg = wait(30000, m.port)
        if type(responseMsg) = "roUrlEvent"
            m.top.response = {
                code: responseMsg.GetResponseCode(),
                content: responseMsg.GetString(),
                context: context
            }
        else
            m.top.response = {code: -1, content: "", context: context}
        end if
    else
        m.top.response = {code: -1, content: "", context: context}
    end if
end sub
