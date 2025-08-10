Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings

Namespace Toml

    Public NotInheritable Class TomlDocument
        Inherits TomlNode

        Public Sub New(rootExpr As TomlExpression)
            MyBase.New(TomlNodeType.Keyvals)

            Dim parentNode As TomlElement = Me

            For Each kvp In rootExpr.Contents
                Select Case kvp.Type
                    Case TomlExpressionType.Keyval
                        ' キーと値のペアを処理
                        Dim keys = ConvertKeys(kvp.Contents(0))
                        parentNode.RegisterKeyValuePair(keys, 0, kvp.Contents(1))

                    Case TomlExpressionType.Table
                        ' テーブルを処理
                        Dim keys = ConvertKeys(kvp.Contents(0))
                        parentNode = Me.RegisterTable(keys, 0, False)

                    Case TomlExpressionType.ArrayTable
                        ' 配列テーブルを処理
                        Dim keys = ConvertKeys(kvp)
                        parentNode = Me.RegisterArrayTable(keys, 0)

                    Case Else
                        ' 他の型は無視
                End Select
            Next
        End Sub

#Region "read"

        ''' <summary>文字列からTOMLドキュメントを読み取ります。</summary>
        ''' <param name="docText">TOMLドキュメント文字列。</param>
        ''' <returns>TOMLドキュメント。</returns>
        Public Shared Function Read(docText As String) As TomlDocument
            ' UTF-8文字列に変換
            Dim u8Text As U8String = U8String.NewString(docText)

            ' TOMLドキュメントを解析
            Dim rootExpr As TomlExpression = TomlParser.Parse(u8Text)
            Return New TomlDocument(rootExpr)
        End Function

        ''' <summary>
        ''' 文字列からTOMLドキュメントを読み取ります。
        ''' 
        ''' このメソッドは、UTF-8文字列を受け取り、TOMLドキュメントを解析して返します。
        ''' 文字列はUTF-8でエンコードされている必要があります。
        ''' </summary>
        ''' <param name="docText">TOMLドキュメント文字列。</param>
        ''' <returns>TOMLドキュメント。</returns>
        Public Shared Function Read(docText As U8String) As TomlDocument
            ' TOMLドキュメントを解析
            Dim rootExpr As TomlExpression = TomlParser.Parse(docText)
            Return New TomlDocument(rootExpr)
        End Function

        Public Shared Function ReadFromFile(docPath As String) As TomlDocument
            Dim path As New IO.FileInfo(docPath)
            If path.Exists Then
                Using rf = New IO.FileStream(docPath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
                    ' BOMを読み取る
                    If rf.Length > 3 Then
                        Dim bom As Byte() = New Byte(2) {}
                        rf.Read(bom, 0, 3)
                        If bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
                            ' UTF-8 BOMが存在する場合、BOMをスキップ
                        Else
                            rf.Seek(0, IO.SeekOrigin.Begin) ' BOMがない場合は先頭に戻す
                        End If
                    End If

                    ' ファイルの内容を読み取る
                    Dim buf As Byte() = New Byte(CInt(path.Length - 1)) {}
                    rf.Read(buf, 0, buf.Length)

                    ' UTF-8文字列に変換
                    Dim docText As U8String = U8String.NewStringChangeOwner(buf)

                    ' TOMLドキュメントを解析
                    Dim rootExpr As TomlExpression = TomlParser.Parse(docText)
                    Return New TomlDocument(rootExpr)
                End Using
            Else
                Throw New IO.FileNotFoundException("TOMLファイルが見つかりません。", docPath)
            End If
        End Function

#End Region

    End Class

End Namespace
