object fmMain: TfmMain
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'EXCEL'#25991#20214#22871#21360#24037#20855
  ClientHeight = 324
  ClientWidth = 644
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -15
  Font.Name = #24494#36575#27491#40657#39636
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  PixelsPerInch = 96
  TextHeight = 19
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
  object cxShellBrowserDialog1: TcxShellBrowserDialog
    FolderLabelCaption = #36984#25799#35201#25918#32622#30003#35531#25991#20214#30340#36039#26009#22846
    LookAndFeel.NativeStyle = True
    Title = #30003#35531#25991#20214#36039#26009#22846
    Left = 548
    Top = 84
  end
  object XLSReadWriteII5: TXLSReadWriteII5
    ComponentVersion = '5.20.67a'
    Version = xvExcel2007
    DirectRead = False
    DirectWrite = False
    Left = 540
    Top = 10
  end
  object cxPropertiesStore1: TcxPropertiesStore
    Components = <
      item
        Component = EditDoc
        Properties.Strings = (
          'Text')
      end
      item
        Component = EditOutputFolder
        Properties.Strings = (
          'Text')
      end
      item
        Component = EditSrc
        Properties.Strings = (
          'Text')
      end>
    StorageName = 'cxPropertiesStore1'
    Left = 595
    Top = 10
  end
end
