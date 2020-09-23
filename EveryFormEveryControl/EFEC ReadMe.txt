 Every Form, Every Controls by Thomas Greenwood
 thomasgreenwood@2die4.com

 No credit Required to use.

 Known Limitations:
   1.) On unload of a form all changes are reset
       Solution: Call SetForms before loading the form.
   2.) My knowledge of the On Error Statement and its non-obvious effects
       are low so if you know, please let me know.

 To prevent controls you do not wish editing from been so, add an If statement to the
 For Each In Next Loops to check the .name of the control (no it won't remain as tempControl),  or set the
 Propertys when the loop(s) is/are finished

 You can also use this statement to process the Type Of Control you wish to edit:
 If TypeOf Control = typeOfControl Then Control.Blah
 e.g. If TypeOf tempControl = TextBox Then tempControl.Text = "TextBox"

 I used this code to set the picture of every form to a dynamically loaded pic

 Thanks Very Much,

 Thomas Greenwood