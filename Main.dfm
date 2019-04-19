object fmMain: TfmMain
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'EXCEL'#25991#20214#22871#21360#24037#20855
  ClientHeight = 351
  ClientWidth = 646
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = #24494#36575#27491#40657#39636
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 19
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 646
    Height = 351
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 649
    ExplicitHeight = 350
    object TabSheet1: TTabSheet
      Caption = #22871#21360
      ExplicitLeft = -36
      ExplicitTop = 100
      ExplicitWidth = 458
      ExplicitHeight = 219
      object Label1: TLabel
        Left = 257
        Top = 241
        Width = 60
        Height = 19
        Caption = #25351#23450#26085#26399
      end
      object EditSrc: TEdit
        Left = 8
        Top = 42
        Width = 618
        Height = 27
        TabOrder = 0
      end
      object btnSrc: TButton
        Left = 8
        Top = 8
        Width = 125
        Height = 28
        Caption = #20358#28304#36039#26009'(Excel)'
        TabOrder = 1
        OnClick = btnSrcClick
      end
      object btnDoc: TButton
        Left = 8
        Top = 84
        Width = 125
        Height = 28
        Caption = #22871#21360#25991#20214'(Excel)'
        TabOrder = 2
        OnClick = btnDocClick
      end
      object EditDoc: TEdit
        Left = 8
        Top = 118
        Width = 618
        Height = 27
        TabOrder = 3
      end
      object btnOuput: TButton
        Left = 8
        Top = 160
        Width = 125
        Height = 28
        Caption = #20786#23384#30446#37636
        TabOrder = 4
        OnClick = btnOuputClick
      end
      object EditOutputFolder: TEdit
        Left = 8
        Top = 194
        Width = 618
        Height = 27
        TabOrder = 5
      end
      object btnExec: TButton
        Left = 8
        Top = 232
        Width = 189
        Height = 37
        Caption = #38283#22987#22871#21360
        TabOrder = 6
        OnClick = btnExecClick
      end
      object ProgressBar1: TProgressBar
        Left = 8
        Top = 275
        Width = 618
        Height = 29
        Step = 1
        TabOrder = 7
      end
      object DateTimePicker_Assign: TDateTimePicker
        Left = 330
        Top = 237
        Width = 121
        Height = 27
        Date = 43454.396880104170000000
        Time = 43454.396880104170000000
        TabOrder = 8
      end
    end
    object TabSheet2: TTabSheet
      Caption = #35722#25976#35498#26126
      ImageIndex = 1
      ExplicitWidth = 281
      ExplicitHeight = 159
      object ListBox1: TListBox
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 632
        Height = 311
        Align = alClient
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -19
        Font.Name = #27161#26999#39636
        Font.Style = []
        ItemHeight = 19
        Items.Strings = (
          '<<'#23458#25142#20195#34399'>>       = '#23458#25142#20195#34399
          '<<'#23458#25142#21517#31281'>>       = '#23458#25142#21517#31281
          '<<'#23458#25142#20840#31281'>>       = '#23458#25142#20840#31281' '
          '<<'#23458#25142#21517#31281'>>       = '#20844#21496#21517#31281
          '<<'#23458#25142#20840#31281'>>       = '#20844#21496#21517#31281
          '<<'#32113#19968#32232#34399'>>       = '#32113#19968#32232#34399
          '<<'#32879#32097#20154'>>         = '#32879#32097#20154' OR '#36899#32097#20154
          '<<'#36899#32097#20154'>>         = '#32879#32097#20154' OR '#36899#32097#20154
          '<<'#22320#22336'>>           = '#22320#22336
          '<<'#22320#22336'>>           = '#20844#21496#22320#22336
          '<<'#38651#35441'>>           = '#38651#35441
          '<<'#38651#35441'>>           = '#32879#32097#38651#35441
          '<<'#38651#35441'>>           = '#36899#32097#38651#35441
          '<<'#20659#30495'>>           = '#20659#30495
          '<<'#35347#32244#24107'>>         = '#35347#32244#24107
          '<<'#36676#21312'>>           = '#36676#21312
          '<<'#20170#22825#26085#26399'>>       = ['#31995#32113#33258#21205#29986#29983']'
          '<<'#25351#23450#26085#26399'>>       = ['#22635#20837#25351#23450#26085#26399']'
          '<<'#24180#26376#26085#27969#27700#34399'>>   = ['#31995#32113#33258#21205#29986#29983']'
          '<<'#20170#22825#26085#26399#27969#27700#34399'>> = ['#31995#32113#33258#21205#29986#29983']'
          '<<'#25351#23450#26085#26399#27969#27700#34399'>> = ['#31995#32113#33258#21205#29986#29983']')
        ParentFont = False
        TabOrder = 0
        ExplicitLeft = 15
        ExplicitTop = 124
        ExplicitWidth = 311
        ExplicitHeight = 287
      end
    end
  end
  object cxShellBrowserDialog1: TcxShellBrowserDialog
    FolderLabelCaption = #36984#25799#35201#25918#32622#30003#35531#25991#20214#30340#36039#26009#22846
    LookAndFeel.NativeStyle = True
    Title = #30003#35531#25991#20214#36039#26009#22846
    Left = 533
    Top = 109
  end
  object XLSReadWriteII5: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 525
    Top = 35
  end
  object cxPropertiesStore1: TcxPropertiesStore
    Components = <>
    StorageName = 'cxPropertiesStore1'
    Left = 580
    Top = 35
  end
  object JcVersionInfo1: TJcVersionInfo
    Left = 315
    Top = 175
  end
end
