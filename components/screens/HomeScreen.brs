sub init()
    m.httpTask = CreateObject("roSGNode", "HttpTask")
    m.httpTask.observeField("response", "onContentLoaded")

    m.top.observeField("authToken", "onAuthToken")

    rowList = m.top.findNode("rowList")
    rowList.observeField("itemSelected", "onItemSelected")
    rowList.observeField("rowItemFocused", "onRowItemFocused")

    if m.top.authToken <> invalid and m.top.authToken <> ""
        loadContent()
    end if
end sub

sub onAuthToken()
    loadContent()
end sub

sub loadContent()
    token = m.top.authToken
    headers = {}
    if token <> invalid and token <> ""
        headers = {Authorization: "Bearer " + token}
    end if

    m.top.findNode("spinner").visible = true
    m.top.findNode("errorLabel").visible = false

    m.httpTask.request = {
        url: "https://reelmotionapp.com/api/browse-data",
        method: "GET",
        headers: headers,
        body: "",
        context: "browseData"
    }
end sub

sub onContentLoaded()
    resp = m.httpTask.response
    m.top.findNode("spinner").visible = false

    if resp = invalid or resp.code <> 200
        m.top.findNode("errorLabel").text = "Could not load content. Please try again."
        m.top.findNode("errorLabel").visible = true
        return
    end if

    json = ParseJson(resp.content)
    if json = invalid
        m.top.findNode("errorLabel").text = "Error parsing content."
        m.top.findNode("errorLabel").visible = true
        return
    end if

    buildRowList(json)
end sub

sub buildRowList(json as Object)
    rowList = m.top.findNode("rowList")

    rootContent = CreateObject("roSGNode", "ContentNode")
    rootContent.CONTENTTYPE = "SECTION"

    ' Continue watching row
    if json.continue_watching <> invalid and json.continue_watching.count() > 0
        rowNode = CreateObject("roSGNode", "ContentNode")
        rowNode.title = "Continue Watching"
        rowNode.CONTENTTYPE = "SECTION"
        for each video in json.continue_watching
            item = makeVideoNode(video)
            rowNode.appendChild(item)
        end for
        rootContent.appendChild(rowNode)
    end if

    ' Category rows
    if json.categories <> invalid
        for each cat in json.categories
            rowNode = CreateObject("roSGNode", "ContentNode")
            rowNode.title = cat.name
            rowNode.CONTENTTYPE = "SECTION"
            if cat.videos <> invalid
                for each video in cat.videos
                    item = makeVideoNode(video)
                    rowNode.appendChild(item)
                end for
            end if
            rootContent.appendChild(rowNode)
        end for
    end if

    rowList.content = rootContent
    rowList.setFocus(true)
end sub

function makeVideoNode(video as Object) as Object
    item = CreateObject("roSGNode", "ContentNode")
    item.title = video.title
    item.HDPosterUrl = video.thumbnail
    item.shortDescriptionLine1 = video.rating
    item.shortDescriptionLine2 = video.duration

    ' Store custom fields
    item.addFields({
        id: "",
        is_free: false
    })
    if video.id <> invalid
        item.id = video.id.toStr()
    end if
    if video.is_free <> invalid
        item.is_free = video.is_free
    end if

    return item
end function

sub onItemSelected()
    rowList = m.top.findNode("rowList")
    selected = rowList.focusedChild
    if selected = invalid then return
    m.top.selectedItem = selected
end sub

sub onRowItemFocused()
    ' Could highlight nav items based on focus position
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if press
        if key = "back"
            m.top.close = true
            return true
        end if
        if key = "up"
            ' When at top of rowList, nothing to do but could focus nav
            return false
        end if
        ' Settings shortcut on options/star key
        if key = "options"
            m.top.goSettings = true
            return true
        end if
    end if
    return false
end function
