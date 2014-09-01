FFT
========

File Filter Toolkit

概要
-----
Windows上でのファイル処理を、定義したフィルタクラスをチェインさせる事によって、UNIXのパイプ処理と同等の事が出来るようにした。 VBScript/WindowsScriptFile実装のフレームワーク

機能一覧
-----
    オプション処理 GetOptArgs, GetOptChoice, GetOptPrompt, GetOptDialog 以降は未実装 GetOptIni, GetOptReg
    ファイル操作   AddFiles, SetExclude, SetExcludeList, SetModifyOnly, SetOrder, SetRename, SetBackup
    フィルタ定義   SetFilter, ClearFilter, DeleteFilter
    組込フィルタ   Grep, Sed, Tr, Sort, Uniq, Empty, Contains 以降は未実装 Cut, Wc, Cat, Split, Tee, Head, Tail
    処理実行、他   Execute, View, Show, Resources, LoadResourceFile
    ユーティリティ M, S, CreateObject, GetTempName, GPN, GFN, GBN, GEN, Split, Sort, Tokenize
                   GetWinFolder, GetSysFolder, GetTempFolder, GetTempFile, LenB, MidB, LeftB, RightB, LPAD, RPAD
                   OpenMDBFile, OpenXMLFile
                   TextFileクラス, IE_IPCクラス

利用例
-----
VbLCmtDel.wsf - VBソースコードの行コメント削除
    
    Call Main
    
    Sub Main
      'ツールキット生成
      Set coFFT = New FileFilterToolkit
      With coFFT
        'wsfファイルに処理ファイルをドロップすると引数展開されるので、それを処理対象ファイルとして登録
        'AddFilesメソッドで、特定ディレクトリ下の全ファイルを再帰的に登録するなどもできます。
        Call .GetOptArgs(WScript.Arguments)
                                            '
                                            
        Call .SetRename("$", ".cln") '処理結果を拡張子を.clnにリネームしたファイルに出力

        Call .SetFilter("VbLCmtDel", New VbLCmtDel) '行コメントを削除するフィルタをセット

        Call .Execute                               'ファイル処理を実行
        'Call .View                                 '処理結果ファイルをメモ帳で表示
      End With
    End Sub
    
    'VBソースコードの行コメント削除 行フィルタオブジェクト
    Class VbLCmtDel
    
      '変数定義
      Public FilterName
      Public FilterType
    
      'フィルタ初期処理
      Public Sub Initialize(argFFT)
        With argFFT
          FilterName = "VbLCmtDel"
          FilterType = "Line"
        End With
      End Sub
    
      'フィルタ終了処理
      Public Sub Terminate(argFFT)
    
      End Sub
    
      'ファイルオープン
      Public Function OpenFile(argFFT, argImpFile, argWrkFile)
        OpenFile = True
      End Function
    
      'ファイルクローズ
      Public Function CloseFile(argFFT, argImpFile, argWrkFile)
        CloseFile = True
      End Function
    
      '行処理
      Public Function ProcessLine(argFFT, argLine)
        If Not argFFT.M(argLine,"^[ \t　]*'", "") Then
          ProcessLine = True
        End If
      End Function
    
    End Class
    
