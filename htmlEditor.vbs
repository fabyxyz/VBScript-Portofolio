Option Explicit
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim objShell
Set objShell = CreateObject("WScript.Shell")
rem version
Dim engineVersion : engineVersion = "v1.0"
rem setup
if not fso.FolderExists("Projects") then
    fso.CreateFolder "Projects"
end if
if not fso.FolderExists("Saves") then
    fso.CreateFolder "Saves"
end if

'///System Variables///'
Dim spl : spl = "HTML Editor"
Dim rt : rt = "Runtime Error"
Dim help : help = "Help"
Dim info : info = "Info"
Dim nw : nw = vbCrlf
Dim nbsp : nbsp = " "
Dim und : und = ""
Dim qt : qt = """"
Dim err : err = 16
Dim crMin, crHr, crDay, crMonth, crYear
crMin = Minute(Now)
crHr = Hour(Now)
crDay = Day(Now)
crMonth = Month(Now)
crYear = Year(Now)
Dim crDate
crDate = crMin & crHr & crDay & crMonth & crYear

'///Global Variables///'
Dim projectName
Call refreshProjectName
Sub refreshProjectName()
    projectName = "project-" & crDate
End Sub

'Basic'
Dim title : title = und
Dim description : description = und
Dim language : language = und
Dim bodyColor : bodyColor = und
Dim chr : chr = und
Dim charset : charset = und
Dim viewport : viewport = false
Dim cssLink : cssLink = und
Dim jsLink : jsLink = und

'Events'
Dim event_resize_px : event_resize_px = "x"
Dim event_resize_color : event_resize_color = "y"
Dim onLoadMessage : onLoadMessage = und

'Content'
rem header
Dim headerHeight : headerHeight = 300
Dim headerColor : headerColor = und
Dim headerGradientSourceColor : headerGradientSourceColor = und
Dim headerGradientDestinationColor : headerGradientDestinationColor = und
Dim hgds : hgds = und
Dim headerGradientDirection : headerGradientDirection = und
Dim hbs : hbs = und
Dim headerBorderColor : headerBorderColor = und
Dim headerBorderThickness : headerBorderThickness = und
Dim headerTitle : headerTitle = und
rem title
Dim titleSize : titleSize = und
Dim titleColor : titleColor = und
Dim isTitleNormal : isTitleNormal = und
Dim isTitleBold : isTitleBold = und
Dim isTitleItalic : isTitleItalic = und
Dim isTitleUnderlined : isTitleUnderlined = und
Dim isTitleOverlined : isTitleOverlined = und
Dim isTitleStrikethrough : isTitleStrikethrough = und
Dim tfs : tfs = und
Dim titleFont : titleFont = und


Call subMain()
Sub subMain()
    'Check time
    Dim time
    time = Hour(Now)
    Dim timeMessage
    if time >= 22 and time <= 23 or time >= 0 and time <= 6 then
        timeMessage = "Good Night"
    elseif time >= 6 and time <= 12 then
        timeMessage = "Good Morning"
    elseif time >= 12 and time <= 17 then
        timeMessage = "Good Afternoon"
    elseif time >= 17 and time <= 22 then
        timeMessage = "Good Evening"
    end if
    Dim main
    main=InputBox("HTML Script Engine" & nbsp & engineVersion & nw & _
                  "====================" & nw & nw & _
                  "1> Create New Project" & nw & _
                  "2> Open Project" & nw & _
                  "3> Run Project" & nw & _
                  "0> Exit" & nw,timeMessage)
        if main = 0 then
            WScript.Quit
        elseif main = 1 then
            refreshProjectName()
            Call subNewProject()
        elseif main = 2 then
            'Call subOpenProject()
            msgBox "Coming Soon...",0+64,"Work In Progress"
                subMain()
        elseif main = 3 then
            'Call subRunProject()
            msgBox "Coming Soon...",0+64,"Work In Progress"
                subMain()
        end if
End Sub

    Sub subNewProject()
        projectName=InputBox("Project Name:","Create New Project",projectName)
            if projectName = "" then
                msgBox "Project Name Required",0+err,rt
                    subNewProject()
            elseif projectName = "?" then
                msgBox "Type in the project name",0+64,help
                    subNewProject()
            else
                Call subProjectSettings()
            end if
    End Sub
        
        Sub subWarnExit()
            Dim warnExit
            warnExit=msgBox("Are you sure you want to exit?" & nw & "You have unsaved changes!",4+32+256+4096,spl)
                if warnExit = vbYes then
                    subMain()
                else
                    Call subProjectSettings()
                end if
        End Sub

        Sub subProjectSettings()
            Dim projectSettings
            projectSettings=InputBox("1> Basic Settings" & nw & _
                                     "2> Advanced Settings" & nw & _
                                     "3> Events" & nw & _
                                     "4> Content" & nw & nw & _
                                     "5> Save Project" & nw & _
                                     "6> Export Project" & nw & _
                                     "0> Exit To Main Menu" & nw,projectName)
                if projectSettings = 0 then
                    subWarnExit()
                elseif projectSettings = "" then
                    subProjectSettings()
                elseif projectSettings = 1 then
                    Call subBasicSettings()
                elseif projectSettings = 2 then
                    Call subAdvancedSettings()
                elseif projectSettings = 3 then
                    Call subEvents()
                elseif projectSettings = 4 then
                    Call subContent()
                elseif projectSettings = 5 then
                    Call subSave()
                elseif projectSettings = 6 then
                    Call subExport()
                end if
        End Sub

            Sub subBasicSettings()
                Dim basicSettings
                basicSettings=InputBox("1> Document Information" & nw & _
                                       "2> Background Color" & nw & _
                                       "0> Return" & nw,"Basic Settings")
                    if basicSettings = 0 then
                        subProjectSettings()
                    elseif basicSettings = 1 then
                        Call subDocInfo()
                    elseif basicSettings = 2 then
                        Call subBackgroundColor()
                    end if
            End Sub

                Sub subDocInfo()
                    Dim docInfo
                    docInfo=InputBox("1> Document Title" & nw & _
                                     "2> Document Description" & nw & _
                                     "3> Document Language" & nw & _
                                     "0> Return" & nw,"Document Information")
                        if docInfo = 0 then
                            subBasicSettings()
                        elseif docInfo = 1 then
                            Call subDocTitle()
                        elseif docInfo = 2 then
                            Call subDocDesc()
                        elseif docInfo = 3 then
                            Call subDocLang()
                        end if
                End Sub

                    Sub subDocTitle()
                        title=InputBox("Title:","Document Title",title)
                        subDocInfo()
                    End Sub

                    Sub subDocDesc()
                        description=InputBox("Description:","Document Description",description)
                        subDocInfo()
                    End Sub

                    Sub subDocLang()
                        language=InputBox("Language:" & nw & "Format Example: " & qt & "en" & qt,"Document Language",language)
                        subDocInfo()
                    End Sub

                Sub subBackgroundColor()
                    bodyColor=InputBox("HEX, RGB or Default (Color Name):","Background (Body) Color",bodyColor)
                    subBasicSettings()
                End Sub

            Sub subAdvancedSettings()
                Dim advancedSettings
                advancedSettings=InputBox("1> Charset" & nw & _
                                          "2> Viewport" & nw & _
                                          "3> Link CSS file" & nw & _
                                          "4> Link JavaScript file" & nw & nw & _
                                          "0> Return"& nw,"Advanced Settings")
                    if advancedSettings = 0 then
                        subProjectSettings()
                    elseif advancedSettings = 1 then
                        Call subCharset()
                    elseif advancedSettings = 2 then
                        Call subViewport()
                    elseif advancedSettings = 3 then
                        Call subLinkCSS()
                    elseif advancedSettings = 4 then
                        Call subLinkJS()
                    end if 
            End Sub

                Sub subCharset()
                    chr=InputBox("1> ASCII" & nw & _
                                 "2> ANSI" & nw & _
                                 "3> 8859" & nw & _
                                 "4> UTF-8" & nw & nw & _
                                 "0> Return" & nw,"Charset",chr)
                        if chr = 0 then
                            subAdvancedSettings()
                                chr = und
                        elseif chr = 1 then
                            charset = "ASCII"
                        elseif chr = 2 then
                            charset = "ANSI"
                        elseif chr = 3 then
                            charset = "8859"
                        elseif chr = 4 then
                            charset = "UTF-8"
                        end if
                            subAdvancedSettings()
                End Sub

                Sub subViewport()
                    Dim vw
                    vw=msgBox("Viewport = " & viewport & nw & _
                              "Yes> True" & nw & _
                              "No> False" & nw,4+32,"Viewport")
                        if vw = vbYes then
                            viewport = true
                        else
                            viewport = false
                        end if
                            subAdvancedSettings()
                End Sub

                Sub subLinkCSS()
                    cssLink=InputBox("CSS File Link/Path:","CSS Stylesheet",cssLink)
                        subAdvancedSettings()
                End Sub

                Sub subLinkJS()
                    jsLink=InputBox("JS File Link/Path:","JavaScript SRC",jsLink)
                        subAdvancedSettings()
                End Sub

            Sub subEvents()
                Call subEvent1()
            End Sub
                
                Sub subEvent1()
                    Dim event1
                    event1=InputBox("1> event_resize" & nw & _
                                    "2> on_load_message" & nw & _
                                    "3> &s" & nw & _
                                    "4> &s" & nw & _
                                    "5> &s" & nw & nw & _
                                    "N> Next Page" & nw & "0> Return","Events - Page 1")
                        if event1 = 0 then
                            subProjectSettings()
                        end if
                        if event1 = "N" then
                            Call subEvent2()
                        end if
                        if event1 = 1 then
                            msgBox "If the browser window is <x> px or smaller, the background will be <y> color.",0+64,info
                            Call subEventResize()
                        elseif event1 = 2 then
                            Call subOnLoadMessage()
                        elseif event1 = 3 then
                            rem event
                        elseif event1 = 4 then
                            rem event
                        elseif event1 = 5 then
                            rem event
                        end if
                End Sub

                    Sub subEventResize()
                            event_resize_px=InputBox("if (px <= " & event_resize_px & ") {" & nw & vbTab & "background-color = " & event_resize_color & ";" & nw & "}" & _
                                nw & nw & "0> Return" & nw & "X:","event_resize (1/2) - Size Required")
                            if event_resize_px = 0 then
                                subEvent1()
                            elseif event_resize_px <= 0  then
                                msgBox "The value cannot be negative!",0+err,rt
                                    event_resize_px = "x"
                                    subEventResize()
                            elseif event_resize_px < 500 and event_resize_px > 0 then
                                msgBox "Value needs to be higher than 500",0+err,rt
                                    event_resize_px = "x"
                                    subEventResize()
                            end if
                        event_resize_color=InputBox("if (px <= " & event_resize_px & ") {" & nw & vbTab & "background-color = " & event_resize_color & ";" & nw & "}" & _
                            nw & nw & "0> Return" & nw & "Y:","event_resize (2/2) - Color Required")
                        subEvent1()
                    End Sub

                    Sub subOnLoadMessage()
                        onLoadMessage=InputBox("Message:","Show alert message when page is loaded",onLoadMessage)
                            subEvent1()
                    End Sub

                Sub subEvent2()
                    Dim event2
                    event2=InputBox("1> &s" & nw & _
                                    "2> &s" & nw & _
                                    "3> &s" & nw & _
                                    "4> &s" & nw & _
                                    "5> &s" & nw & nw & _
                                    "N> Next Page" & nw & "P> Previous Page" & nw & "0> Return","Events - Page 2")
                        rem if
                End Sub

            Sub subContent()
                Call subContent1()
            End Sub

                Sub subContent1()
                    Dim content1
                    content1=InputBox("1> Header" & nw & _
                                      "2> Page Title" & nw & _
                                      "3> Add Paragraph" & nw & _
                                      "4> &s" & nw & _
                                      "5> &s" & nw & nw & _
                                      "N> Next Page" & nw & "0> Return" & nw,"Content - Page 1")
                        if content1 = 0 then
                            subProjectSettings()
                        end if
                        if content1 = "N" then
                            Call subContent2()
                        end if
                        if content1 = 1 then
                            Call subHeader()
                        elseif content1 = 2 then
                            Call subTitle()
                        elseif content1 = 3 then
                            Call subPr()
                        elseif content1 = 4 then
                            rem null
                        elseif content1 = 5 then
                            rem null
                        end if
                End Sub

                    Sub subHeader()
                        Dim headerSettings
                        headerSettings=InputBox("1> Header Height" & nw & _
                                                "2> Header Color" & nw & _
                                                "3> Header Border" & nw & nw & _
                                                "0> Return" & nw,"Header Settings")
                            if headerSettings = 0 then
                                subContent1()
                            end if
                            if headerSettings = 1 then
                                Call subHeaderHeight()
                            elseif headerSettings = 2 then
                                Call subHeaderColorSettings()
                            elseif headerSettings = 3 then
                                Call subHeaderBorderSettings()
                            end if
                    End Sub

                        Sub subHeaderHeight()
                            headerHeight=InputBox("PX:","Header Height",headerHeight)
                                if headerHeight < 0 then
                                    msgBox "The value cannot be negative!",0+err,rt
                                        subHeaderHeight()
                                else
                                    subHeader()
                                end if
                        End Sub

                        Sub subHeaderColorSettings()
                            Dim headerColorSettings
                            headerColorSettings=InputBox("1> Set Color" & nw & "2> Set Gradient" & nw & nw & "0> Return" & nw,"Header Color Settings")
                                if headerColorSettings = 0 then
                                    subHeader()
                                end if
                                if headerColorSettings = 1 then
                                    Call subHeaderColor()
                                elseif headerColorSettings = 2 then
                                    Call subHeaderGradient()
                                end if
                        End Sub

                            Sub subHeaderColor()
                                headerColor=InputBox("HEX, RGB or Default (Color Name):","Header Color",headerColor)
                                    subHeaderColorSettings()
                            End Sub

                            Sub subHeaderGradient()
                                Dim headerGradient
                                headerGradient=InputBox("1> Color" & nw & "2> Direction" & nw & nw & "0> Return" & nw,"Header Gradient Settings")
                                    if headerGradient = 1 then
                                        Call subHeaderGradientColor()
                                    elseif headerGradient = 2 then
                                        Call subHeaderGradientDirection()
                                    else
                                        Call subHeaderColorSettings()
                                    end if
                            End Sub

                                Sub subHeaderGradientColor()
                                    headerGradientSourceColor=InputBox("HEX, RGB or Default (Color Name):","Header Gradient Source Color",headerGradientSourceColor)
                                    headerGradientDestinationColor=InputBox("HEX, RGB or Default (Color Name):","Header Gradient Destination Color",headerGradientDestinationColor)
                                        subHeaderGradient()
                                End Sub

                                Sub subHeaderGradientDirection()
                                    hgds=InputBox("1> To Top" & nw & _
                                                  "2> To Bottom" & nw & _
                                                  "3> To Right" & nw & _
                                                  "4> To Left" & nw & nw & _
                                                  "0> Return" & nw,"Header Gradient Direction")
                                        if hgds = 0 then
                                            subHeaderGradient()
                                        end if
                                        if hgds = 1 then
                                            headerGradientDirection = "to top"
                                        elseif hgds = 2 then
                                            headerGradientDirection = "to bottom"
                                        elseif hgds = 3 then
                                            headerGradientDirection = "to right"
                                        elseif hgds = 4 then
                                            headerGradientDirection = "to left"
                                        end if
                                            subHeaderGradient()
                                End Sub

                        Sub subHeaderBorderSettings()
                            hbs=InputBox("1> Border Color" & nw & "2> Border Thickness" & nw & nw & "0> Return" & nw,"Header Border Settings")
                                if hbs = 0 then
                                    subHeader()
                                end if
                                if hbs = 1 then
                                    Call subHeaderBorderColor()
                                elseif hbs = 2 then
                                    Call subHeaderBorderThickness()
                                end if
                        End Sub

                            Sub subHeaderBorderColor()
                                headerBorderColor=InputBox("HEX, RGB or Default (Color Name):","Header Border Color",headerBorderColor)
                                    subHeaderBorderSettings()
                            End Sub

                            Sub subHeaderBorderThickness()
                                headerBorderThickness=InputBox("Thickness (px):","Header Border Thickness",headerBorderThickness)
                                    subHeaderBorderSettings()
                            End Sub

                    Sub subTitle()
                        Dim titleSettings
                        titleSettings=InputBox("1> Set Title" & nw & _
                                               "2> Size" & nw & _
                                               "3> Color" & nw & _
                                               "4> Style" & nw & _
                                               "5> Font Family" & nw & nw & _
                                               "0> Return" & nw,"Title Settings")
                            if titleSettings = 0 then
                                subContent1()
                            end if
                            if titleSettings = 1 then
                                headerTitle=InputBox("Title:","Page Title",headerTitle)
                                    subTitle()
                            elseif titleSettings = 2 then
                                titleSize=InputBox("Size (px):","Title Size",titleSize)
                                    subTitle()
                            elseif titleSettings = 3 then
                                titleColor=InputBox("HEX, RGB or Default (Color Name):","Title Color",titleColor)
                                    subTitle()
                            elseif titleSettings = 4 then
                                Call subTitleStyle()
                            elseif titleSettings = 5 then
                                Call subTitleFont()
                            end if
                    End Sub

                        Sub subTitleStyle()
                            Dim tts
                            Dim ttsinp(6)
                            Dim ts : ts = "Title Style"
                            tts=InputBox("1> Normal" & nw & _
                                         "2> Bold" & nw & _
                                         "3> Italic" & nw & _
                                         "4> Underline" & nw & _
                                         "5> Overline" & nw & _
                                         "6> Strikethrough" & nw & nw & _
                                         "0> Return" & nw,"Header Title Style")
                                if tts = 0 then
                                    subTitle()
                                end if
                                if tts = 1 then
                                    ttsinp(0)=msgBox("Is Title Normal?",3+32,ts)
                                        if ttsinp(0) = vbYes then
                                            isTitleNormal = true
                                                subTitleStyle()
                                        elseif ttsinp(0) = vbNo then
                                            isTitleNormal = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                elseif tts = 2 then
                                    ttsinp(1)=msgBox("Is Title Bold?",3+32,ts)
                                        if ttsinp(1) = vbYes then
                                            isTitleBold = true
                                                subTitleStyle()
                                        elseif ttsinp(1) = vbNo then
                                            isTitleBold = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                elseif tts = 3 then
                                    ttsinp(2)=msgBox("Is Title Italic?",3+32,ts)
                                        if ttsinp(2) = vbYes then
                                            isTitleItalic = true
                                                subTitleStyle()
                                        elseif ttsinp(2) = vbNo then
                                            isTitleItalic = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                elseif tts = 4 then
                                    ttsinp(3)=msgBox("Is Title Underlined?",3+32,ts)
                                        if ttsinp(3) = vbYes then
                                            isTitleUnderlined = true
                                                subTitleStyle()
                                        elseif ttsinp(3) = vbNo then
                                            isTitleUnderlined = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                elseif tts = 5 then
                                    ttsinp(4)=msgBox("Is Title Overlined?",3+32,ts)
                                        if ttsinp(4) = vbYes then
                                            isTitleOverlined = true
                                                subTitleStyle()
                                        elseif ttsinp(4) = vbNo then
                                            isTitleOverlined = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                elseif tts = 6 then
                                    ttsinp(5)=msgBox("Is Title Strikethrough?",3+32,ts)
                                        if ttsinp(5) = vbYes then
                                            isTitleStrikethrough = true
                                                subTitleStyle()
                                        elseif ttsinp(5) = vbNo then
                                            isTitleStrikethrough = false
                                                subTitleStyle()
                                        else
                                            subTitleStyle()
                                        end if
                                end if
                        End Sub

                        Sub subTitleFont()
                            Dim tf
                            tf=InputBox("Font:" & nw & "1> Clear Font" & nw & "0> Return" & nw,"Header Title Font",titleFont)
                                if tf = 0 then
                                    subTitle()
                                elseif tf = 1 then
                                    titleFont = und
                                        subTitle()
                                else
                                    titleFont = tf
                                    tf = und
                                        subTitle()
                                end if
                        End Sub

                    Sub subPr()
                        
                    End Sub

                Sub subContent2()
                    Dim content2
                    content2=InputBox("1> &s" & nw & _
                                      "2> &s" & nw & _
                                      "3> &s" & nw & _
                                      "4> &s" & nw & _
                                      "5> &s" & nw & nw & _
                                      "P> Previous Page" & nw & "0> Return","Content - Page 2")
                        if content2 = 0 then
                            subProjectSettings()
                        end if
                        if content2 = "P" then
                            subContent1()
                        end if
                End Sub

            Sub subSave()
                rem null
            End Sub

            Sub subExport()
                if fso.FileExists("Projects/" & projectName & ".html") then
                    Dim prex
                    prex=msgBox("A project with this name already exists" & nw & _
                           "Yes> Rename Current Project" & nw & _
                           "No> Overwrite Project" & nw & nw & _
                           "Cancel> Return" & nw,3+err,spl)
                        if prex = vbYes then
                            Call renameProject()
                        elseif prex = vbNo then
                            fso.DeleteFile "Projects/" & projectName & ".html"
                                Call export()
                        elseif prex = vbCancel then
                            subProjectSettings()
                        end if
                else
                    Call export()
                end if
            End Sub

                Sub renameProject()
                    projectName=InputBox("Project Name:","Rename Project",projectName)
                        if projectName = "" then
                            msgBox "Project Name Required",0+err,spl
                                renameProject()
                        elseif projectName = "?" then
                            msgBox "Type in a new project name",0+64,help
                                renameProject()
                        else
                            Call subExport()
                        end if
                End Sub

                Sub export()
                    Dim html
                    Set html = fso.CreateTextFile("Projects\" & projectName & ".html")
                    WScript.Sleep 1000
                    html.WriteLine "<!DOCTYPE html>"
                    if language <> und then
                        html.WriteLine "<html lang=" & qt & language & qt & ">"
                    else
                        html.WriteLine "<html>"
                    end if
                    rem head
                    html.WriteLine "<head>"
                    if title <> und then
                        html.WriteLine vbTab & "<title>" & title & "</title>"
                    end if
                    if description <> und then
                        html.WriteLine vbTab & "<meta name=" & qt & "description" & qt & nbsp & "content=" & qt & description & qt & ">"
                    end if
                    if charset <> und then
                        html.WriteLine vbTab & "<meta charset=" & qt & charset & qt & ">"
                    end if
                    if viewport = true then
                        html.WriteLine vbTab & "<meta name=" & qt & "viewport" & qt & nbsp & "content=" & _
                            qt & "width=device-width, initial-scale=1.0" & qt & ">"
                    end if
                    if cssLink <> und then
                        html.WriteLine vbTab & "<link rel=" & qt & "stylesheet" & qt & nbsp & "href=" & qt & cssLink & qt & ">"
                    end if
                    if jsLink <> und then
                        html.WriteLine vbTab & "<script src=" & qt & jsLink & qt & "></script>"
                    end if
                    '///
                    if onLoadMessage <> und then
                        html.WriteLine "<script>"
                        html.WriteLine vbTab & "function onLoadMessage() {"
                        html.WriteLine vbTab & vbTab & "alert(" & qt & onLoadMessage & qt & ");"
                        html.WriteLine vbTab & "}"
                        html.WriteLine "</script>"
                    end if
                    html.WriteLine "</head>"
                    html.WriteBlankLines(1)
                    '///
                    html.Write "<body"
                    if bodyColor <> und then
                        html.Write " bgcolor=" & qt & bodyColor & qt
                    end if
                    if onLoadMessage <> und then
                        html.Write " onload=" & qt & "onLoadMessage()" & qt
                    end if
                    html.Write ">"
                    rem Body
                    html.WriteBlankLines(1)
                    if headerHeight <> und then
                        html.WriteLine vbTab & "<div class=" & qt & "header" & qt & ">"
                    end if
                    if headerTitle <> und then
                        html.WriteLine vbTab & "<p class=" & qt & "headerTitle" & qt & ">" & headerTitle & "</p>"
                    end if
                    '//end
                    html.WriteLine "</body>"
                    html.WriteBlankLines(1)
                    html.WriteLine "</html>"

                    html.WriteBlankLines(1)

                    html.WriteLine "<style>"
                    rem style
                    if event_resize_px <> "x" and event_resize_color <> "y" then
                        html.WriteLine vbTab & "@media only screen and (max-width: " & event_resize_px & "px) {"
                        html.WriteLine vbTab & vbTab & "body {"
                        html.WriteLine vbTab & vbTab & vbTab & "background-color: " & event_resize_color & ";"
                        html.WriteLine vbTab & vbTab & "}"
                        html.WriteLine vbTab & "}"
                    end if
                    'header style
                    if headerHeight <> und then
                        html.WriteLine vbTab & ".header {"
                        html.WriteLine vbTab & vbTab & "margin-left: auto;"
                        html.WriteLine vbTab & vbTab & "margin-right: auto;"
                        html.WriteLine vbTab & vbTab & "height: " & headerHeight & "px;"
                        if headerBorderColor <> und and headerBorderThickness <> und then
                            html.WriteLine vbTab & vbTab & "width: calc(99% - " & headerBorderThickness * 2 & "px);"
                        else
                            html.WriteLine vbTab & vbTab & "width: 100%;"
                        end if
                        if headerColor <> und then
                            html.WriteLine vbTab & vbTab & "background-color: " & headerColor & ";"
                        end if
                        if headerGradientDirection <> und then
                            html.WriteLine vbTab & vbTab & "background-image: linear-gradient(" & headerGradientDirection & ", " & _
                                headerGradientSourceColor & ", " & headerGradientDestinationColor & ");"
                        end if
                        if headerBorderColor <> und and headerBorderThickness <> und then
                            html.WriteLine vbTab & vbTab & "border: " & headerBorderThickness & "px solid" & nbsp & headerBorderColor & ";"
                        end if
                        html.WriteLine vbTab & "}"
                    end if
                    if headerTitle <> und then
                        html.WriteLine vbTab & ".headerTitle {"
                            html.WriteLine vbTab & vbTab & "text-align: center;"
                            html.WriteLine vbTab & vbTab & "margin-top:" & nbsp & headerHeight / 2 & "px;"
                            if titleSize <> und then
                                html.WriteLine vbTab & vbTab & "font-size:" & nbsp & titleSize & "px;"
                            end if
                            if titleColor <> und then
                                html.WriteLine vbTab & vbTab & "color:" & nbsp & titleColor & ";"
                            end if
                            if isTitleNormal = true then
                                html.WriteLine vbTab & vbTab & "font-style: normal;"
                            end if
                            if isTitleBold = true then
                                html.WriteLine vbTab & vbTab & "font-weight: bold;"
                            end if
                            if isTitleItalic = true then
                                html.WriteLine vbTab & vbTab & "font-style: italic;"
                            end if
                            if isTitleUnderlined = true then
                                html.WriteLine vbTab & vbTab & "text-decoration: underline;"
                            end if
                            if isTitleOverlined = true then
                                html.WriteLine vbTab & vbTab & "text-decoration: overline;"
                            end if
                            if isTitleStrikethrough = true then
                                html.WriteLine vbTab & vbTab & "text-decoration: line-through;"
                            end if
                        html.WriteLine vbTab & "}"
                    end if

                    html.WriteLine "</style>"

                    html.WriteBlankLines(1)

                    html.WriteLine "<script>"
                    '///
                    html.WriteLine "</script>"

                    rem source watermark
                    html.WriteBlankLines(1)
                    html.WriteLine "<!--Made With HTML Script Engine" & nbsp & engineVersion & nbsp & "-->"
                    html.Write "<!--@fabyxyz-->"

                    html.Close
                    
                    if fso.FileExists("Projects/" & projectName & ".html") then
                        msgBox "Find your projects in:" & nw & "(Projects/" & projectName & ".html)",0+64,"Project Succefully Exported!"
                    else
                        msgBox "The project could not be exported",0+err,spl
                    end if
                End Sub
