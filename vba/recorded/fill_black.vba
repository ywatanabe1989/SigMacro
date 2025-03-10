Option Explicit
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
ActiveDocument.CurrentPageItem.Select(False, -2036, 1734, -2036, 1734)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00000000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H00000000&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
End Sub