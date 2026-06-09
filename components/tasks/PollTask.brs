sub init()
    m.top.functionName = "runPoll"
    m.top.sessionToken = ""
    m.top.activationComplete = false
    m.top.codeExpired = false
    m.top.logMessage = "PollTask: initialized"
end sub

sub runPoll()
    deviceToken = m.top.deviceToken
    baseUrl = "https://reelmotionapp.com"

    if deviceToken = invalid or deviceToken = ""
        logDebug("PollTask: missing device token; cannot start polling")
        return
    end if

    logDebug("PollTask: started polling")

    while true
        ' Wait 5 seconds between polls.
        sleep(5000)

        response = postPollRequest(baseUrl, deviceToken)
        responseCode = response.code
        responseBody = response.body

        logDebug("PollTask: poll response code = " + str(responseCode).trim())
        logDebug("PollTask: poll response body = " + truncateForLog(responseBody, 160))

        if responseBody <> ""
            json = ParseJson(responseBody)
            if json <> invalid and json.status <> invalid
                if json.status = "activated"
                    sessionToken = json.session_token
                    if sessionToken = invalid then sessionToken = ""

                    if sessionToken <> ""
                        sec = CreateObject("roRegistrySection", "reelmotion")
                        sec.Write("session_token", sessionToken)
                        sec.Flush()
                        logDebug("PollTask: session token saved to registry")
                    else
                        logDebug("PollTask: activated response did not include session_token")
                    end if

                    m.top.sessionToken = sessionToken
                    m.top.activationComplete = true
                    return
                else if json.status = "expired"
                    logDebug("PollTask: activation code expired")
                    m.top.codeExpired = true
                    return
                else
                    logDebug("PollTask: activation status = " + json.status)
                end if
            else
                logDebug("PollTask: could not parse poll response JSON")
            end if
        else
            logDebug("PollTask: empty response; polling will retry")
        end if
    end while
end sub

function postPollRequest(baseUrl as String, deviceToken as String) as Object
    port = CreateObject("roMessagePort")
    url = CreateObject("roUrlTransfer")
    url.SetMessagePort(port)
    url.SetUrl(baseUrl + "/api/auth/device/poll")
    url.SetCertificatesFile("common:/certs/ca-bundle.crt")
    url.InitClientCertificates()
    url.EnableEncodings(true)
    url.RetainBodyOnError(true)
    url.SetConnectTimeout(30000)
    url.AddHeader("Content-Type", "application/json")
    url.AddHeader("Accept", "application/json")
    url.SetRequest("POST")

    body = FormatJson({ device_token: deviceToken })
    logDebug("PollTask: posting to " + baseUrl + "/api/auth/device/poll")

    if url.AsyncPostFromString(body) <> true
        logDebug("PollTask: AsyncPostFromString failed to start: " + url.GetFailureReason())
        return makePollResponse(-1, "")
    end if

    urlEvent = wait(30000, port)
    if urlEvent = invalid
        url.AsyncCancel()
        logDebug("PollTask: poll request timed out")
        return makePollResponse(-1, "")
    end if

    if type(urlEvent) <> "roUrlEvent"
        logDebug("PollTask: unexpected poll event type = " + type(urlEvent))
        return makePollResponse(-1, "")
    end if

    responseBody = urlEvent.GetString()
    if responseBody = invalid then responseBody = ""

    failureReason = urlEvent.GetFailureReason()
    if failureReason <> invalid and failureReason <> ""
        logDebug("PollTask: poll failure reason = " + failureReason)
    end if

    return makePollResponse(urlEvent.GetResponseCode(), responseBody)
end function

function makePollResponse(responseCode as Integer, responseBody as String) as Object
    response = {}
    response.code = responseCode
    response.body = responseBody
    return response
end function

sub logDebug(message as String)
    print message
    m.top.logMessage = message
end sub

function truncateForLog(value as String, maxLength as Integer) as String
    if value = invalid then return ""
    if len(value) <= maxLength then return value
    return left(value, maxLength) + "..."
end function
