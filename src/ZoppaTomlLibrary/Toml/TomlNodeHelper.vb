Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>
    ''' TOMLノードのヘルパーモジュール。
    ''' </summary>
    ''' <remarks>
    ''' このモジュールは、TOMLノードの操作に関連するヘルパー関数を提供します。
    ''' </remarks>
    Module TomlNodeHelper

        ''' <summary>キーをドットで区切って、UTF-8文字列のリストに変換します。</summary>
        ''' <param name="keyExpr">キーを表す式。</param>
        ''' <returns>UTF-8文字列のリスト。</returns>
        Function ConvertKeys(keyExpr As TomlExpression) As List(Of U8String)
            Dim keys As New List(Of U8String)()
            For Each subKeyExpr In keyExpr.Contents
                Select Case subKeyExpr.Type
                    Case TomlExpressionType.UnquotedKey
                        ' 引用符なしのキー
                        keys.Add(subKeyExpr.Str)

                    Case TomlExpressionType.QuotedKey
                        ' 引用符付きのキー
                        Dim keyStr = subKeyExpr.Contents(0).Str
                        Select Case subKeyExpr.Contents(0).Type
                            Case TomlExpressionType.BasicString
                                keys.Add(ConvertToBasicString(keyStr.GetIterator()))
                            Case TomlExpressionType.LiteralString
                                keys.Add(ConvertToLiteralString(keyStr.GetIterator()))
                            Case Else
                                Throw New TomlSyntaxException($"不正なキーの型: {subKeyExpr.Type}")
                        End Select

                    Case Else
                        Throw New TomlSyntaxException($"不正なキーの型: {subKeyExpr.Type}")
                End Select
            Next
            Return keys
        End Function

        ''' <summary>インラインテーブルを作成します。</summary>
        ''' <param name="inlineExpr">インラインテーブルを表す式。</param>
        ''' <returns>登録されたインラインテーブルノード。</returns>
        ''' <remarks>インラインテーブルは、通常の値ツリーではなく、ノードツリーに登録されます。</remarks>
        Function CreateInlineTable(inlineExpr As TomlExpression) As TomlNode
            Dim rootNode = New TomlInlineTable()

            For Each kvp In inlineExpr.Contents
                Select Case kvp.Type
                    Case TomlExpressionType.Keyval
                        ' キーと値のペアを処理
                        Dim keys = ConvertKeys(kvp.Contents(0))
                        rootNode.RegisterKeyValuePair(keys, 0, kvp.Contents(1))

                    Case TomlExpressionType.Table
                        ' テーブルを処理
                        Dim keys = ConvertKeys(kvp.Contents(0))
                        rootNode.RegisterTable(keys, 0, False)

                    Case Else
                        ' 他の型は無視
                        Throw New TomlSyntaxException($"不正なインラインテーブルの値: {kvp.Str}")
                End Select
            Next

            Return rootNode
        End Function

        ''' <summary>配列を作成します。</summary>
        ''' <param name="arrayExpr">配列を表す式。</param>
        ''' <returns>登録された配列ノード。</returns>
        ''' <remarks>配列は、通常の値ツリーではなく、ノードツリーに登録されます。</remarks>
        Function CreateArray(arrayExpr As TomlExpression) As TomlElement
            Dim resNode As New TomlArray()
            For Each item In arrayExpr.Contents
                Select Case item.Type
                    Case TomlExpressionType.InlineTable
                        ' インラインテーブルの場合、ノードツリーに登録
                        resNode.AddItem(CreateInlineTable(item))

                    Case TomlExpressionType.Array
                        ' 配列の場合、配列ノードを作成してノードツリーに登録
                        resNode.AddItem(CreateArray(item))

                    Case Else
                        ' 通常の値の場合、値ツリーに登録
                        resNode.AddItem(item)
                End Select
            Next
            Return resNode
        End Function

    End Module

End Namespace
