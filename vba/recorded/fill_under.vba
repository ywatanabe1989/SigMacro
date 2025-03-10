Option Explicit
Function FlagOn(flag As Long)
  FlagOn = flag Or FLAG_SET_BIT ' Use to set option flag bits on, leaving others unchanged
End Function
Function FlagOff(flag As Long)
  FlagOff = flag Or FLAG_CLEAR_BIT ' Use to set option flag bits off, leaving others unchanged
End Function
Sub Main
ActiveDocument.CurrentPageItem.Select(False, -2072, 1734, -2072, 1734)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00800080&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00800080&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLOR, &H00800080&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SEA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H0000ff00&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLOR, &H0000ff00&)
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, SDA_COLORREPEAT, &H00000002&)
ActiveDocument.CurrentPageItem.Select(False, 2058, 1135, 2058, 1135)
ActiveDocument.CurrentPageItem.Select(False, -1562, 1692, -1562, 1692)
ActiveDocument.CurrentPageItem.GraphPages(0).CurrentPageObject(GPT_GRAPH).NameObject.SetObjectCurrent
ActiveDocument.CurrentPageItem.SetCurrentObjectAttribute(GPM_SETPLOTATTR, &H000008a7&, &H0000001e&)
ActiveDocument.CurrentPageItem.Select(False, 761, 1651, 761, 1651)
End Sub