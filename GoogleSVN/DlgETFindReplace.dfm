object FindReplaceDlg: TFindReplaceDlg
  Left = 260
  Top = 231
  BorderIcons = [biSystemMenu]
  Caption = 'Find and Replace'
  ClientHeight = 425
  ClientWidth = 451
  Color = clBtnFace
  Constraints.MinWidth = 459
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poOwnerFormCenter
  OnCreate = FormCreate
  OnDeactivate = FormDeactivate
  OnKeyDown = FormKeyDown
  OnResize = FormResize
  DesignSize = (
    451
    425)
  PixelsPerInch = 96
  TextHeight = 13
  object gpResizer: TGrip
    Left = 435
    Top = 409
    Width = 16
    Height = 16
    DrawHandler = False
  end
  object tcFindReplace: TTntTabControl
    Left = 7
    Top = 7
    Width = 437
    Height = 161
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 0
    Tabs.Strings = (
      'Fin&d'
      'Re&place'
      '&Go to')
    TabIndex = 0
    OnChange = tcFindReplaceChange
    OnEnter = tcFindReplaceEnter
    DesignSize = (
      437
      161)
    object lblFind: TTntLabel
      Left = 8
      Top = 34
      Caption = 'Fi&nd what:'
      FocusControl = mcbFind
    end
    object lblReplace: TTntLabel
      Left = 8
      Top = 57
      Caption = 'R&eplace with:'
      FocusControl = mcbReplace
    end
    object lblWithin: TTntLabel
      Left = 8
      Top = 88
      Caption = 'Wit&hin:'
      FocusControl = cbRange
    end
    object lblSearch: TTntLabel
      Left = 8
      Top = 112
      Caption = '&Search:'
      FocusControl = cbOrder
    end
    object lblLookin: TTntLabel
      Left = 8
      Top = 136
      Caption = '&Look in:'
      FocusControl = cbLookin
    end
    object mcbFind: TMRUComboBox
      Left = 80
      Top = 30
      Width = 347
      AutoComplete = False
      Anchors = [akLeft, akTop, akRight]
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 0
      OnChange = mcbFindChange
    end
    object mcbReplace: TMRUComboBox
      Left = 80
      Top = 54
      Width = 347
      AutoComplete = False
      Anchors = [akLeft, akTop, akRight]
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      TabOrder = 1
      OnChange = mcbReplaceChange
    end
    object cbRange: TTntComboBox
      Left = 80
      Top = 84
      Width = 86
      AutoComplete = False
      Style = csDropDownList
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      ItemIndex = 0
      TabOrder = 2
      Text = 'Sheet'
      OnSelect = cbRangeSelect
      Items.Strings = (
        'Sheet'
        'Workbook')
    end
    object cbOrder: TTntComboBox
      Left = 80
      Top = 108
      Width = 86
      AutoComplete = False
      Style = csDropDownList
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      ItemIndex = 0
      TabOrder = 3
      Text = 'By Rows'
      OnSelect = cbOrderSelect
      Items.Strings = (
        'By Rows'
        'By Columns')
    end
    object cbLookin: TTntComboBox
      Left = 80
      Top = 132
      Width = 86
      AutoComplete = False
      Style = csDropDownList
      ImeName = #20013#25991' ('#31616#20307') - '#25628#29399#25340#38899#36755#20837#27861
      ItemIndex = 0
      TabOrder = 4
      Text = 'Formulas'
      OnSelect = cbLookinSelect
      Items.Strings = (
        'Formulas'
        'Values'
        'Comments')
    end
    object cbMatchCase: TTntCheckBox
      Left = 175
      Top = 83
      Caption = 'Match &case'
      TabOrder = 5
      OnClick = cbMatchCaseClick
    end
    object cbMatchContents: TTntCheckBox
      Left = 175
      Top = 102
      Caption = 'Match entire cell c&ontents'
      TabOrder = 6
      OnClick = cbMatchContentsClick
    end
    object cbMatchByte: TTntCheckBox
      Left = 175
      Top = 121
      Caption = 'Match &byte'
      TabOrder = 7
      OnClick = cbMatchByteClick
    end
    object btnOptions: TTntButton
      Left = 354
      Top = 129
      Anchors = [akRight, akBottom]
      Caption = 'Op&tions <<'
      TabOrder = 8
      OnClick = btnOptionsClick
    end
  end
  object tcGoto: TTntTabControl
    Left = 7
    Top = 8
    Width = 437
    Height = 161
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 1
    Tabs.Strings = (
      'Fin&d'
      'Re&place'
      '&Go to')
    TabIndex = 0
    OnChange = tcFindReplaceChange
    DesignSize = (
      437
      161)
    object ghSelect: TKSOGroupHeader
      Left = 8
      Top = 24
      Width = 421
      Caption = 'Select'
    end
    object lblDataType: TTntLabel
      Left = 32
      Top = 86
      Caption = 'Data Type:'
    end
    object rbData: TTntRadioButton
      Left = 16
      Top = 42
      Caption = 'D&ata'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = rbDataClick
    end
    object rbComments: TTntRadioButton
      Tag = 1
      Left = 16
      Top = 110
      Caption = '&Comments'
      TabOrder = 7
      OnClick = rbDataClick
    end
    object rbBlanks: TTntRadioButton
      Tag = 2
      Left = 168
      Top = 110
      Caption = 'Blan&ks'
      TabOrder = 8
      OnClick = rbDataClick
    end
    object rbLastCell: TTntRadioButton
      Tag = 4
      Left = 16
      Top = 134
      Caption = 'La&st Cell'
      TabOrder = 10
      OnClick = rbDataClick
    end
    object rbVisibleCellsOnly: TTntRadioButton
      Tag = 3
      Left = 288
      Top = 110
      Caption = 'Visible Cells Onl&y'
      TabOrder = 9
      OnClick = rbDataClick
    end
    object cbConstants: TTntCheckBox
      Left = 32
      Top = 62
      Caption = 'C&onstants'
      TabOrder = 1
      OnClick = cbConstantsClick
    end
    object cbFormulas: TTntCheckBox
      Left = 128
      Top = 62
      Caption = 'Fo&rmulas'
      TabOrder = 2
      OnClick = cbConstantsClick
    end
    object cbErrors: TTntCheckBox
      Left = 314
      Top = 86
      Caption = '&Errors'
      TabOrder = 6
      OnClick = cbNumbersClick
    end
    object cbText: TTntCheckBox
      Left = 169
      Top = 86
      Caption = 'Te&xt'
      TabOrder = 4
      OnClick = cbNumbersClick
    end
    object cbLogicals: TTntCheckBox
      Left = 233
      Top = 86
      Caption = '&Logicals'
      TabOrder = 5
      OnClick = cbNumbersClick
    end
    object cbNumbers: TTntCheckBox
      Left = 96
      Top = 86
      Caption = 'N&umbers'
      TabOrder = 3
      OnClick = cbNumbersClick
    end
    object rbHyperlink: TTntRadioButton
      Tag = 6
      Left = 344
      Top = 46
      Caption = 'Hyperl&ink'
      TabOrder = 11
      Visible = False
      OnClick = rbDataClick
    end
    object rbCurDataArea: TTntRadioButton
      Tag = 5
      Left = 168
      Top = 134
      Caption = 'Curre&nt Data Area'
      TabOrder = 12
      OnClick = rbDataClick
    end
    object rbObject: TTntRadioButton
      Tag = 7
      Left = 288
      Top = 136
      Caption = 'O&bject'
      TabOrder = 13
      OnClick = rbDataClick
    end
  end
  object btnGoto: TTntButton
    Left = 286
    Top = 174
    Anchors = [akTop, akRight]
    Caption = '&Go to'
    Default = True
    TabOrder = 6
    OnClick = btnGotoClick
  end
  object btnReplaceAll: TTntButton
    Left = 7
    Top = 174
    Anchors = [akTop, akRight]
    Caption = 'Replace &All'
    TabOrder = 2
    OnClick = btnReplaceAllClick
    OnEnter = OnEnterSearch
    OnExit = OnExitSearch
  end
  object btnReplace: TTntButton
    Left = 89
    Top = 174
    Anchors = [akTop, akRight]
    Caption = '&Replace'
    TabOrder = 3
    OnClick = btnReplaceClick
    OnEnter = OnEnterSearch
    OnExit = OnExitSearch
  end
  object btnFindAll: TTntButton
    Left = 171
    Top = 174
    Anchors = [akTop, akRight]
    Caption = 'F&ind All'
    TabOrder = 4
    OnClick = btnFindAllClick
    OnEnter = OnEnterSearch
    OnExit = OnExitSearch
  end
  object btnFindNext: TTntButton
    Left = 277
    Top = 174
    Width = 85
    Anchors = [akTop, akRight]
    Caption = '&Find Next'
    Default = True
    TabOrder = 5
    OnClick = btnFindNextClick
    OnEnter = OnEnterSearch
    OnExit = OnExitSearch
  end
  object btnClose: TTntButton
    Left = 369
    Top = 174
    Anchors = [akTop, akRight]
    Cancel = True
    Caption = 'Close'
    TabOrder = 7
    OnClick = btnCloseClick
  end
  object lvFindAll: TTntListView
    Left = 7
    Top = 205
    Width = 437
    Height = 212
    Anchors = [akLeft, akTop, akRight, akBottom]
    Columns = <
      item
        Caption = 'Book'
        Width = 70
      end
      item
        Caption = 'Sheet'
        Width = 70
      end
      item
        Caption = 'Name'
      end
      item
        Caption = 'Cell'
        Width = 70
      end
      item
        Caption = 'Value'
      end
      item
        AutoSize = True
        Caption = 'Formula'
      end>
    MultiSelect = True
    ReadOnly = True
    RowSelect = True
    ShowWorkAreas = True
    TabOrder = 8
    ViewStyle = vsReport
    OnClick = OnClick
    OnColumnClick = OnColClick
    OnCompare = lvFindAllCompare
    OnResize = lvFindAllResize
    OnSelectItem = OnSelectItem
  end
end
