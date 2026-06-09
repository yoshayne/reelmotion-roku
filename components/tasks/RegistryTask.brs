sub init()
    m.port = CreateObject("roMessagePort")
    m.top.observeField("request", m.port)
end sub

sub mainThread()
    while true
        msg = wait(0, m.port)
        if type(msg) = "roSGNodeEvent"
            if msg.getField() = "request"
                regInput()
            end if
        end if
    end while
end sub

sub regInput()
    req = m.top.request
    if req = invalid then return
    context = req.context
    if context = invalid then return

    parameters = context.parameters
    if parameters = invalid then return

    command = parameters.command
    section = parameters.section
    key = parameters.key

    if command = "read"
        val = RegRead(key, section)
        context.response = {regVal: val}
    else if command = "write"
        RegWrite(key, parameters.value, section)
        context.response = {regVal: "ok"}
    else if command = "delete"
        RegDelete(key, section)
        context.response = {regVal: "ok"}
    else if command = "deleteSection"
        DeleteSection(section)
        context.response = {regVal: "ok"}
    end if
end sub

function RegRead(key as String, section as String) as String
    if section = invalid or section = "" then section = "Default"
    sec = CreateObject("roRegistrySection", section)
    if sec.Exists(key)
        return sec.Read(key)
    end if
    return "invalid"
end function

sub RegWrite(key as String, val as String, section as String)
    if section = invalid or section = "" then section = "Default"
    sec = CreateObject("roRegistrySection", section)
    sec.Write(key, val)
    sec.Flush()
end sub

sub RegDelete(key as String, section as String)
    if section = invalid or section = "" then section = "Default"
    sec = CreateObject("roRegistrySection", section)
    sec.Delete(key)
    sec.Flush()
end sub

sub DeleteSection(section as String)
    if section = invalid or section = "" then return
    reg = CreateObject("roRegistry")
    reg.Delete(section)
    reg.Flush()
end sub
