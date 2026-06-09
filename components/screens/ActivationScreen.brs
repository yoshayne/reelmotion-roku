sub init()
    ' Create a simple task to test HTTP connectivity
    m.testTask = CreateObject("roSGNode", "TestHttpTask")
    m.testTask.observeField("result", "onTestResult")
    m.testTask.control = "RUN"

    ' Show loading text
    m.top.findNode("codeLabel").text = "Testing connection..."
end sub

sub onTestResult()
    m.top.findNode("codeLabel").text = m.testTask.result
end sub
