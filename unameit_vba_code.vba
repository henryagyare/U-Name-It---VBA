Dim button_clicked As String

Sub Beans_Click()
    Dim result As Integer
    button_clicked = "beans"
    MsgBox "Click OK to learn some facts about beans", vbOKCancel, UCase(button_clicked)
    MsgBox "Beans are an excellent source of vegetable protein and minerals such as iron, magnesium and zinc", vbInformation, UCase(button_clicked)
    MsgBox "Beans are rich in folic acid, an element associated with the reduction of such birth defects as Spina Bifida", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show the Beans image.", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Greens_Click()
    Dim result As Integer
    button_clicked = "greens"
    MsgBox "Prepare to learn some facts about Greens", vbDefaultButton1, UCase(button_clicked)
    MsgBox "Dark green vegetables are rich in vitamins A, C, and K, and folate.", vbInformation, UCase(button_clicked)
    MsgBox "These nutrients protect bones, decrease inflammation, help with vision, improve immunity, and protect against some types of cancers.", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Potatoes_Click()
    Dim result As Integer
    button_clicked = "potatoes"
    MsgBox "Did you Know?" & vbCrLf & _
    " " & vbCrLf & _
    "Today potatoes are grown in all 50 states of the USA and in about 125 countries throughout the world.", vbInformation, UCase(button_clicked)
    MsgBox "An 8 ounce baked or boiled potato has only about 100 calories", vbInformation, UCase(button_clicked)
    MsgBox "Thomas Jefferson is credited for introducing “french fries” to America when he served them at a White House dinner.", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Tomatoes_Click()
    Dim result As Integer
    Dim yesNo As Integer
    
    button_clicked = "tomatoes"
    yesNo = MsgBox("Are you ready to learn some interesting things about Tomatoes?", vbYesNo, UCase(button_clicked))
    If yesNo = vbYes Then
        MsgBox "Tomatoes are the major dietary source of the antioxidant lycopene" & vbCrLf & _
        "Lycopene has been linked to many health benefits, including reduced risk of heart disease and cancer.", vbInformation, UCase(button_clicked)
        MsgBox "Tomatoes are considered beneficial for skin health." & vbCrLf & _
        " " & vbCrLf & _
        "Tomato-based foods rich in lycopene and other plant compounds may protect against sunburn", vbInformation, UCase(button_clicked)
        result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
        ' Check if user clicked OK
        If result = vbOK Then
            ' Show the image on the slide
            ShowImagePopup
        End If
    End If
End Sub

Sub Lambs_Click()
    Dim result As Integer
    Dim yesNo As Integer
    button_clicked = "lambs"
        
    yesNo = MsgBox("Are you ready to learn about Lambs?", vbYesNo, UCase(button_clicked))
    If yesNo = vbYes Then
        MsgBox "A Lamb is a young sheep" & vbCrLf & _
        " " & vbCrLf & _
        "They are very smart and gentle creatures with an excellent memory." & vbCrLf & _
        " " & vbCrLf & _
        "They build friendships with other sheep and can recognize up to 50 other sheep faces as well as remember human faces.", vbInformation, UCase(button_clicked)
        result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
        ' Check if user clicked OK
        If result = vbOK Then
            ' Show the image on the slide
            ShowImagePopup
        End If
    End If
End Sub

Sub Rams_Click()
    Dim result As Integer
    button_clicked = "rams"
    MsgBox "Here is some information about Rams: " & vbCrLf & _
    " " & vbCrLf & _
    "Rams are male bighorn sheep, animals that live in the mountains" & vbCrLf & _
    " " & vbCrLf & _
    "and often settle arguments with fights that include ramming their heads into others", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Hogs_Click()
    Dim result As Integer
    button_clicked = "hogs"
    MsgBox "Here is some information about Hogs: " & vbCrLf & _
    " " & vbCrLf & _
    "The largest wild pig is the giant forest hog, which can grow up to 6.6 feet long." & vbCrLf & _
    "And the heaviest is the Eurasian wild pig, which can reach a whopping 710 pounds.", vbInformation, UCase(button_clicked)
    
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Dogs_Click()
    Dim result As Integer
    Dim yesNo As Integer
    button_clicked = "dogs"
    yesNo = MsgBox("Are you ready to learn some interesting facts about Dogs?", vbYesNo, UCase(button_clicked))
    MsgBox "A dog’s nose print is unique, much like a person’s fingerprint.", vbInformation, UCase(button_clicked)
    MsgBox "A dog’s sense of smell is legendary, but did you know that their nose has as many as 300 million receptors? ", vbInformation, UCase(button_clicked)
    MsgBox "In comparison, a human nose has about 5 million", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Chickens_Click()
    Dim result As Integer
    button_clicked = "chickens"
    MsgBox "Be Prepared to learn some awesome facts about Chikens!", vbInformation, UCase(button_clicked)
    MsgBox "Chickens are the closest living relatives of dinosaurs!", vbInformation, UCase(button_clicked)
    MsgBox "Scientific evidence has proven the shared common ancestry between chickens and the Tyrannosaurus rex.", vbInformation, UCase(button_clicked)
    MsgBox "Chickens have over 30 unique vocalizations that they use to communicate a wide variety of messages to other chickens, ", vbInformation, UCase(button_clicked)
    MsgBox "These vocalizations Include: " & vbCrLf & _
    " " & vbCrLf & _
    "mating calls, stress signals, warnings of danger, how they are feeling and food discovery.", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Turkeys_Click()
    Dim result As Integer
    button_clicked = "turkeys"
    MsgBox "Click OK to learn some fun facts about Turkeys!", vbInformation, UCase(button_clicked)
    MsgBox "Turkey droppings tell a bird’s sex and age. Male droppings are j-shaped; female droppings are spiral-shaped.", vbInformation, UCase(button_clicked)
    MsgBox "The larger the diameter, the older the bird.", vbInformation, UCase(button_clicked)
    MsgBox "Feathers galore: An adult turkey has 5,000 to 6,000 feathers", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub Rabbits_Click()
    Dim result As Integer
    button_clicked = "rabbits"
    MsgBox "Brace for some fun facts about Rabbits!", vbInformation, UCase(button_clicked)
    MsgBox " A baby rabbit is called a kit, a female is called a doe and a male is called a buck.", vbInformation, UCase(button_clicked)
    MsgBox "A rabbit’s teeth never stop growing! Instead, they’re gradually worn down as the rabbit chews on grasses, wildflowers and vegetables", vbInformation, UCase(button_clicked)
    MsgBox "meaning they never get too long", vbInformation, UCase(button_clicked)
    result = MsgBox("Click OK to show an image of " + button_clicked + ".", vbOKCancel, UCase(button_clicked))
    ' Check if user clicked OK
    If result = vbOK Then
        ' Show the image on the slide
        ShowImagePopup
    End If
End Sub

Sub ShowImagePopup()
    Dim oSlide As Slide
    Dim oRect As Shape
    Dim oPicture As Shape
    Dim oButton As Shape
    Dim imagePath As String
    
    ' Set the path to the image
    imagePath = "\unameitImages\" + button_clicked + ".jpg"
    
    ' Get the active slide
    Set oSlide = ActivePresentation.Slides(1)

    ' Insert image into slide
    Set oPicture = oSlide.Shapes.AddPicture(imagePath, msoFalse, msoCTrue, 280, 160, 400, 300)
    
    ' Add a close button (an X) at the top right corner
        ' To add a shape: Set oButton = oSlide.Shapes.AddShape(Type of shape(example msoShapeOval), left margin(how far from the left), top margin, shape width, height)
    Set oButton = oSlide.Shapes.AddShape(msoShapeOval, 880, 170, 20, 20)
        oButton.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Red button
        oButton.TextFrame.TextRange.Text = "X"
    
    ' Assign an action to the button to call ClosePopup when clicked
    With oButton.ActionSettings(ppMouseClick)
        .Action = ppActionRunMacro
        .Run = "ClosePopup"
    End With
End Sub

Sub ClosePopup()
    Dim oSlide As Slide
    Set oSlide = ActivePresentation.Slides(1)

    ' Loop through shapes and delete the rectangle, picture, and close button
    Dim oShp As Shape
    For Each oShp In oSlide.Shapes
        If oShp.Type = msoPicture Or oShp.Type = msoShapeOval Then
            oShp.Delete
        End If
        If oShp.Type = msoShapeOval Then
            oShp.Delete
        End If
    Next oShp
End Sub
