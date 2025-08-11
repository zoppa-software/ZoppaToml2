Option Strict On
Option Explicit On

Imports System.Runtime.CompilerServices
Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>
    ''' TOMLの式を表す構造体。
    ''' この構造体は、TOMLの式の種類とその文字列を保持します。
    ''' </summary>
    ''' <param name="Type">式の種類を示す列挙型。</param>
    ''' <param name="Str">式を表す文字列。</param>
    Public Module TomlParser

        ''' <summary>
        ''' TOMLの式を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' この関数は、TOMLの式を解析し、結果をTomlExpressionとして返します。
        ''' </remarks>
        Public Function Parse(source As U8String) As TomlExpression
            Dim iter = source.GetIterator()

            ' 解析結果を格納するリスト
            Dim expressions As New List(Of TomlExpression)()
            While iter.HasNext()
                ' 式を解析
                Dim expr = ParseExpression(source, iter)
                If expr.Type <> TomlExpressionType.None Then
                    expressions.Add(expr)
                End If

                ' 改行が来た場合は改行を読み捨てる
                If Not MatchNewline(iter) Then
                    Exit While
                End If
            End While

            ' 末尾の 0x00 文字を読み捨てる
            While iter.HasNext() AndAlso iter.Current?.Raw0 = 0
                iter.MoveNext()
            End While

            ' 解析が完了した後にまだ文字が残っている場合は、例外をスロー
            If iter.HasNext() Then
                Throw New TomlParseException("未解析な文字が残っています。", source, iter)
            End If

            ' 解析結果をTomlExpressionとして返す
            Return New TomlExpression(TomlExpressionType.Expression, U8String.Empty, expressions.ToArray())
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、TOMLの式を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' この関数は、TOMLの式を解析し、結果をTomlExpressionとして返します。
        ''' 空白やコメントも含めて解析します。
        ''' </remarks>
        Public Function ParseExpression(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            ' 空白読み捨て
            ParseWs(source, iter)

            ' キー値ペアが取れたら取得、取れなかったらテーブルを取得
            Dim exper = ParseKeyval(source, iter)
            If exper.Type = TomlExpressionType.None Then
                exper = ParseTable(source, iter)
            End If

            ' 取得できたら空白を読み捨てる
            If exper.Type <> TomlExpressionType.None Then
                ParseWs(source, iter)
            End If

            ' コメント読み捨て
            ParseComment(source, iter)

            Return exper
        End Function

#Region "空白"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、空白文字を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 空白文字はスペース（0x20）やタブ（0x09）などを含みます。
        ''' </remarks>
        Public Function ParseWs(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 対象文字であることを確認
            ' 1. 空白文字（スペースやタブ）を読み捨てる
            ' 2. 空白文字が来ない場合は終了
            While iter.HasNext()
                If AnyCharByte(iter.Current, &H20, &H9) Then  ' 1
                    iter.MoveNext()
                Else
                    Exit While ' 2
                End If
            End While

            ' 解析結果をTomlExpressionとして返す
            If iter.CurrentIndex > startIndex Then
                Return New TomlExpression(TomlExpressionType.Ws, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                Return TomlExpression.Empty
            End If
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから空白文字を読み捨てます。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>空白文字が読み捨てられた場合はTrue、それ以外はFalse。</returns>
        Private Function MatchWs(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 対象文字であることを確認
            ' 1. 空白文字（スペースやタブ）を読み捨てる
            ' 2. 空白文字が来ない場合は終了
            While iter.HasNext()
                If AnyCharByte(iter.Current, &H20, &H9) Then  ' 1
                    iter.MoveNext()
                Else
                    Exit While ' 2
                End If
            End While

            ' 解析結果を返す
            Return iter.CurrentIndex > startIndex
        End Function

#End Region

#Region "改行"

        ''' <summary>
        ''' 引数で指定されたU8Charから改行文字をチェックします。
        ''' 改行文字はLF（0x0A）またはCRLF（0x0D 0x0A）です。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>改行文字ならば真。</returns>
        Private Function MatchNewline(iter As U8String.U8StringIterator) As Boolean
            If iter.Current?.Raw0 = &HA Then
                ' LF (0x0A) の場合
                iter.MoveNext()
                Return True

            ElseIf iter.Current?.Raw0 = &HD AndAlso iter.Peek(1)?.Raw0 = &HA Then
                ' CR (0x0D) 、LF (0x0A) の場合
                iter.MoveNext()
                iter.MoveNext()
                Return True
            End If

            Return False
        End Function

#End Region

#Region "コメント"

        ''' <summary>
        ''' 引数の文字が非ASCII文字であるかをチェックします。
        ''' 非ASCII文字は、Unicodeの範囲で0x80以上の文字を指します。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>非ASCII文字ならば真。</returns>
        Private Function IsNonAscii(u8c As U8Char?) As Boolean
            If u8c.HasValue Then
                Return (u8c.Value.IntegerValue >= &H80)
            End If
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Charが非EOL（End of Line）文字であるかをチェックします。
        ''' 非EOL文字は、スペース、タブ、非ASCII文字などを含みます。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>非EOL文字ならば真。</returns>
        Private Function IsNonEol(u8c As U8Char?) As Boolean
            If u8c.HasValue Then
                Return (u8c.Value.Raw0 = &H9 OrElse
                        (u8c.Value.Raw0 >= &H20 AndAlso u8c.Value.Raw0 <= &H7F) OrElse
                        IsNonAscii(u8c))
            End If
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringからコメントを解析します。
        ''' コメントは#で始まり、行末までの文字列を含みます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' コメントは#で始まり、行末までの文字列を含みます。
        ''' </remarks>
        Public Function ParseComment(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If EqualCharByte(iter.Current, &H23) Then ' #
                iter.MoveNext()

                ' コメントの解析
                ' 1. 非EOL文字が来る場合はコメントの一部として読み進める
                ' 2. 改行文字が出現した場合は終了
                While iter.HasNext()
                    If IsNonEol(iter.Current) Then
                        iter.MoveNext() ' 1
                    Else
                        Exit While ' 2
                    End If
                End While

                ' 解析結果をTomlExpressionとして返す
                Return New TomlExpression(TomlExpressionType.Comment, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                ' コメント以外の文字が来た場合は空の式を返す
                Return TomlExpression.Empty
            End If
        End Function

#End Region

#Region "キー、値ペア"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、キーと値のペアを解析します。
        ''' キーと値は、キー=値の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseKeyval(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' キーを取得
            Dim key = ParseKey(source, iter)
            If key.Type = TomlExpressionType.None Then
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If

            ' キーと値の区切り（=）を取得
            Dim keyvalSep = ParseKeyvalOrDotSep(source, iter, &H3D, TomlExpressionType.KeyvalSep)
            If keyvalSep.Type = TomlExpressionType.None Then
                iter.SetCurrentIndex(startIndex)
                Throw New TomlParseException("キーと値の区切りがありません。", source, iter)
            End If

            ' 値を取得
            Dim value = ParseVal(source, iter)
            If value.Type <> TomlExpressionType.None Then
                Return New TomlExpression(TomlExpressionType.Keyval, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex), New TomlExpression() {key, value})
            Else
                ' 値が解析できなかった場合は、イテレータの位置を元に戻して例外をスロー
                iter.SetCurrentIndex(startIndex)
                Throw New TomlParseException("値がありません。", source, iter)
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、キーを解析します。
        ''' キーは単純なキーまたはドット区切りのキーとして解析されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' キーはTOMLの構文で使用される文字列を表します。
        ''' </remarks>
        Function ParseKey(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' キーを解析
            Dim sKey = ParseSimpleKey(source, iter)
            If sKey.Type = TomlExpressionType.None Then
                ' キーが解析できなかった場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If

            ' キーが解析できた場合は、ドット区切りの式を追加していく
            Dim expressions As New List(Of TomlExpression) From {
                sKey
            }

            Do
                ' ドット区切りが来た場合は、次のキーを解析
                Dim dotExpr = ParseKeyvalOrDotSep(source, iter, &H2E, TomlExpressionType.DotSep)
                If dotExpr.Type = TomlExpressionType.DotSep Then
                    sKey = ParseSimpleKey(source, iter)
                    If sKey.Type <> TomlExpressionType.None Then
                        expressions.Add(sKey)
                    Else
                        ' 次のキーが解析できなかった場合は終了
                        Throw New TomlParseException("無効なキーの形式です。", source, iter)
                    End If
                Else
                    ' ドット区切りが来ない場合は終了
                    Exit Do
                End If
            Loop While iter.HasNext()

            ' 解析結果をTomlExpressionとして返す
            Return New TomlExpression(
                If(expressions.Count <= 1, TomlExpressionType.SimpleKey, TomlExpressionType.DottedKey),
                U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex),
                expressions.ToArray()
            )
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、単純なキーを解析します。
        ''' 単純なキーは引用符なしのキーまたは引用符付きのキーとして解析されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 単純なキーは、TOMLのキーとして使用される文字列を表します。
        ''' </remarks>
        Public Function ParseSimpleKey(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 引用符のキーを解析
            Dim quotedExpr = ParseQuotedKey(source, iter)
            If quotedExpr.Type = TomlExpressionType.QuotedKey Then
                Return quotedExpr
            End If

            ' 引用符なしのキーを解析
            Dim unquotedKey = ParseUnquotedKey(source, iter)
            If unquotedKey.Type = TomlExpressionType.UnquotedKey Then
                Return unquotedKey
            End If

            ' どちらの形式も解析できなかった場合は、イテレータの位置を元に戻して空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、引用符なしのキーを解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseUnquotedKey(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 対象文字であることを確認
            While iter.HasNext()
                If IsAlpha(iter.Current) OrElse
                   RangeCharByte(iter.Current, &H30, &H39) OrElse ' 0-9の範囲
                   AnyCharByte(iter.Current, &H2D, &H5F) Then ' - または _
                    iter.MoveNext()
                Else
                    ' 不正な文字が出現した場合は終了
                    Exit While
                End If
            End While

            ' 解析結果をTomlExpressionとして返す
            If iter.CurrentIndex > startIndex Then
                Return New TomlExpression(TomlExpressionType.UnquotedKey, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                ' キーが空の場合は空の式を返す
                Return TomlExpression.Empty
            End If
        End Function

        ''' <summary>引数で指定されたU8Stringから、区切りの式を解析します。</summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="sepByte">区切り文字。</param>
        ''' <param name="exprType">式のタイプ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Private Function ParseKeyvalOrDotSep(source As U8String, iter As U8String.U8StringIterator, sepByte As Byte, exprType As TomlExpressionType) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 空白読み捨て
            ParseWs(source, iter)

            ' 区切り文字判定
            ' 1. 区切り文字が来た場合はを読み進める
            ' 2. 区切り文字が来ない場合はイテレータの位置を元に戻して空の式を返す
            If EqualCharByte(iter.Current, sepByte) Then ' .
                iter.MoveNext() ' 1
            Else
                iter.SetCurrentIndex(startIndex)    ' 2
                Return TomlExpression.Empty
            End If

            ' 空白読み捨て
            ParseWs(source, iter)

            ' 解析結果をTomlExpressionとして返す
            Return New TomlExpression(exprType, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、引用符付きのキーを解析します。
        ''' 引用符付きのキーは、基本文字列またはリテラル文字列として解析されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseQuotedKey(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 基本文字列が解析できた場合は、キーとして返す
            Dim bstrExpr = ParseBasicString(source, iter)
            If bstrExpr.Type = TomlExpressionType.BasicString Then
                Return New TomlExpression(TomlExpressionType.QuotedKey, bstrExpr.Str, New TomlExpression() {bstrExpr})
            End If

            ' リテラル文字列が解析できた場合は、キーとして返す
            Dim lstrExpr = ParseLiteralString(source, iter)
            If lstrExpr.Type = TomlExpressionType.LiteralString Then
                Return New TomlExpression(TomlExpressionType.QuotedKey, lstrExpr.Str, New TomlExpression() {lstrExpr})
            End If

            ' どちらの形式も解析できなかった場合は、イテレータの位置を元に戻して空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、値を解析します。
        ''' 値は基本文字列、真偽値、配列、インラインテーブル、日付、実数、整数などを含みます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseVal(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 文字列
            Dim stringExpr = ParseString(source, iter)
            If stringExpr.Type <> TomlExpressionType.None Then
                Return stringExpr
            End If

            ' 真偽値
            Dim booleanExpr = ParseBoolean(source, iter)
            If booleanExpr.Type <> TomlExpressionType.None Then
                Return booleanExpr
            End If

            ' 配列
            Dim arrayExpr = ParseArray(source, iter)
            If arrayExpr.Type <> TomlExpressionType.None Then
                Return arrayExpr
            End If

            ' インラインテーブル
            Dim inlineTableExpr = ParseInlineTable(source, iter)
            If inlineTableExpr.Type <> TomlExpressionType.None Then
                Return inlineTableExpr
            End If

            ' 日付
            Dim dateExpr = ParseDateTime(source, iter)
            If dateExpr.Type <> TomlExpressionType.None Then
                Return dateExpr
            End If

            ' 実数
            Dim floatExpr = ParseFloat(source, iter)
            If floatExpr.Type <> TomlExpressionType.None Then
                Return floatExpr
            End If

            ' 整数
            Dim intExpr = ParseInteger(source, iter)
            If intExpr.Type <> TomlExpressionType.None Then
                Return intExpr
            End If

            ' 値がない場合は、イテレータの位置を元に戻して空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

#End Region

#Region "文字列"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、文字列を解析します。
        ''' 文字列は基本文字列またはリテラル文字列として解析されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseString(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 複数行文字列を解析
            Dim mlBasicStringExpr = ParseMlBasicString(source, iter)
            If mlBasicStringExpr.Type <> TomlExpressionType.None Then
                Return mlBasicStringExpr
            End If

            ' 基本文字列を解析
            Dim basicStringExpr = ParseBasicString(source, iter)
            If basicStringExpr.Type <> TomlExpressionType.None Then
                Return basicStringExpr
            End If

            ' 複数行リテラル文字列を解析
            Dim mlLiteralStringExpr = ParseMlLiteralString(source, iter)
            If mlLiteralStringExpr.Type <> TomlExpressionType.None Then
                Return mlLiteralStringExpr
            End If

            ' リテラル文字列を解析
            Dim literalStringExpr = ParseLiteralString(source, iter)
            If literalStringExpr.Type <> TomlExpressionType.None Then
                Return literalStringExpr
            End If

            ' 値がない場合は、イテレータの位置を元に戻して空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' 基本文字列
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、基本文字列を解析します。
        ''' 基本文字列はダブルコーテーション（"）で囲まれた文字列です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseBasicString(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If EqualCharByte(iter.Current, &H22) Then ' "
                iter.MoveNext()
                While iter.HasNext()
                    ' 基本文字列でなければ終了
                    If Not MatchBasicChar(iter) Then
                        Exit While
                    End If
                End While

                ' 解析結果をTomlExpressionとして返す
                If EqualCharByte(iter.Current, &H22) Then ' "
                    iter.MoveNext()
                    Return New TomlExpression(TomlExpressionType.BasicString, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 基本文字列の終了記号がない場合は例外をスロー
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("文字列の終了記号がありません。", source, iter)
                End If
            End If

            ' 基本文字列の開始記号がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、基本文字を解析します。
        ''' 基本文字は基本文字列の中で使用される文字で、エスケープされていないものとエスケープされたものがあります。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>基本文字ならば真。</returns>
        Function MatchBasicChar(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 基本文字列のエスケープを判定
            ' 1. 非エスケープの基本文字であるかをチェック
            ' 2. エスケープ文字であるかをチェック
            If IsBasicUnescaped(iter.Current) Then
                iter.MoveNext()             ' 1
                Return True
            ElseIf MatchEscaped(iter) Then  ' 2
                Return True
            End If

            ' 基本文字でなければ偽
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Charが基本文字列のエスケープされていない文字であるかをチェックします。
        ''' 基本文字列では、空白文字や特定の記号が許可されています。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>基本文字列のエスケープされていない文字ならば真。</returns>
        Private Function IsBasicUnescaped(u8c As U8Char?) As Boolean
            Return AnyCharByte(u8c, &H20, &H9, &H21) OrElse
                   RangeCharByte(u8c, &H23, &H5B) OrElse
                   RangeCharByte(u8c, &H5D, &H7E) OrElse
                   IsNonAscii(u8c)
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、エスケープ文字を解析します。
        ''' エスケープ文字はバックスラッシュ（\）で始まり、特定の文字が続きます。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>エスケープ文字ならば真。</returns>
        Function MatchEscaped(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' エスケープ文字の開始を確認、エスケープ文字ならば真を返す
            If EqualCharByte(iter.Current, &H5C) Then ' \
                iter.MoveNext()
                If MatchEscapeSeqChar(iter) Then
                    Return True
                End If
            End If

            ' エスケープ文字でなければ偽
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Charがエスケープシーケンスの文字であるかをチェックします。
        ''' エスケープシーケンスは、バックスラッシュ（\）で始まり、特定の文字が続きます。
        ''' </summary>
        ''' <param name="u8c1">前のU8Char。</param>
        ''' <param name="u8c2">現在のU8Char。</param>
        ''' <returns>エスケープシーケンスの文字ならば真。</returns>
        Private Function IsEscapedSeqChar(u8c1 As U8Char?, u8c2 As U8Char?) As Boolean
            Return EqualCharByte(u8c1, &H5C) AndAlso
                   AnyCharByte(u8c2, &H5C, &H22, &H62, &H66, &H6E, &H72, &H74, &H75, &H55)
        End Function

        ''' <summary>エスケープが連続しているかチェックします。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>エスケープが連続している場合はTrue、それ以外はFalse。</returns>
        Private Function MatchEscapeSeqChar(iter As U8String.U8StringIterator) As Boolean
            ' エスケープシーケンスの文字をチェック
            If iter.Current.HasValue Then
                Select Case iter.Current.Value.Raw0
                    Case &H22, &H5C, &H62, &H66, &H6E, &H72, &H74
                        iter.MoveNext()
                        Return True
                    Case &H75
                        iter.MoveNext()
                        Return MatchHexdigTimes(iter, 4)
                    Case &H55
                        iter.MoveNext()
                        Return MatchHexdigTimes(iter, 8)
                End Select
            End If
            Return False
        End Function

        '------------------------------
        ' 複数行基本文字列
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、複数行基本文字列を解析します。
        ''' 複数行基本文字列は3つのダブルコーテーション（"""）で囲まれた文字列です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseMlBasicString(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If MatchMlStringDelim(iter, &H22) Then
                ' 改行文字を読み捨てる
                MatchNewline(iter)

                ' 複数行基本文字列の内容を読み取る
                If Not MatchMlBasicBody(iter) Then
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("複数行基本文字列の内容がありません。", source, iter)
                End If

                ' 解析結果をTomlExpressionとして返す
                If MatchMlStringDelim(iter, &H22) Then
                    Return New TomlExpression(TomlExpressionType.MlBasicString, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 複数行基本文字列の終了記号がない場合は例外をスロー
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("複数行基本文字列の終了記号がありません。", source, iter)
                End If
            End If

            ' 複数行基本文字列の開始記号がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 複数行基本文字列の内容を解析します。
        ''' ここでは、複数行基本文字列の本体部分を読み取ります。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行基本文字列の内容が正しく解析できた場合は真。</returns>
        Private Function MatchMlBasicBody(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 複数行基本文字列の内容を読み取る
            While MatchMlbContent(iter)
                ' 連続して複数行基本文字列の内容がマッチする限りループ
            End While

            ' 複数行基本文字列の内容を読み進める
            While iter.HasNext()
                Dim enable = False
                Dim midIndex = iter.CurrentIndex
                If MatchMlQuotes(iter, &H22) Then
                    While MatchMlbContent(iter)
                        enable = True
                    End While
                End If
                If Not enable Then
                    iter.SetCurrentIndex(midIndex)
                    Exit While
                End If
            End While

            ' 最後にダブルコーテーションが連続しているかをチェック
            Dim lastIndex = iter.CurrentIndex
            While iter.HasNext() AndAlso EqualCharByte(iter.Current, &H22)
                iter.MoveNext()
            End While

            ' 最後のダブルコーテーションの判定
            Dim num = iter.CurrentIndex - lastIndex
            If num >= 3 AndAlso num <= 5 Then
                iter.SetCurrentIndex(lastIndex + (num - 3))
                Return True
            Else
                iter.SetCurrentIndex(lastIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 複数行基本文字列内の文字を解析します。
        ''' ここでは、複数行基本文字列の内容を読み取ります。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行基本文字列の内容が正しく解析できた場合は真。</returns>
        Private Function MatchMlbContent(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            If MatchMlbChar(iter) Then
                ' 複数行基本文字列内の文字がマッチした場合
                Return True
            ElseIf MatchNewline(iter) Then
                ' 改行がマッチした場合
                Return True
            ElseIf MatchMlbEscapedNl(iter) Then
                ' エスケープされた改行がマッチした場合
                Return True
            End If

            ' 複数行基本文字列内の文字がマッチしなかった場合は、イテレータの位置を元に戻す
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、複数行基本文字列内の文字をチェックします。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行基本文字列内の文字ならば真。</returns>
        Private Function MatchMlbChar(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            If IsMlbUnescaped(iter.Current) Then
                ' エスケープされていない文字の場合
                iter.MoveNext()
                Return True
            ElseIf MatchEscaped(iter) Then
                ' エスケープされた改行の場合
                Return True
            End If

            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>
        ''' 文字列のエスケープされていない文字をチェックします。
        ''' 基本文字列や複数行基本文字列の中で使用される文字がエスケープされていないかを確認します。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>エスケープされていない文字の場合はTrue、それ以外はFalse。</returns>
        Private Function IsMlbUnescaped(u8c? As U8Char) As Boolean
            Return AnyCharByte(u8c, &H20, &H9, &H21) OrElse
                   RangeCharByte(u8c, &H23, &H5B) OrElse
                   RangeCharByte(u8c, &H5D, &H7E) OrElse
                   IsNonAscii(u8c)
        End Function

        ''' <summary>
        ''' 複数行基本文字列の中で改行を読み捨てるための関数です。
        ''' エスケープされた改行も考慮して、改行を正しく処理します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>改行が読み捨てられた場合は真。</returns>
        Private Function MatchMlbEscapedNl(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' エスケープ文字が来た場合は、改行を読み捨てる
            If EqualCharByte(iter.Current, &H5C) Then ' \
                iter.MoveNext()

                ' 空白文字を読み捨てる
                If AnyCharByte(iter.Current, &H20, &H9) Then
                    iter.MoveNext()
                End If

                ' 改行文字を読み捨てる
                If MatchNewline(iter) Then
                    ' 空白と改行を読み捨てる
                    While iter.HasNext()
                        If AnyCharByte(iter.Current, &H20, &H9) Then
                            iter.MoveNext()
                        ElseIf MatchNewline(iter) Then
                            ' 改行を読み捨てる
                        Else
                            ' 改行以外の文字が出現した場合は終了
                            Exit While
                        End If
                    End While

                    ' 改行が読み捨てられた場合は真を返す
                    Return True
                End If
            End If

            ' 複数行基本文字列内に改行がなければ偽をかえす
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        '------------------------------
        ' リテラル文字列
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、リテラル文字列を解析します。
        ''' リテラル文字列はアポストロフィ（'）で囲まれた文字列です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseLiteralString(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If EqualCharByte(iter.Current, &H27) Then ' '
                iter.MoveNext()
                While iter.HasNext()
                    If IsLiteralChar(iter.Current) Then
                        ' リテラル文字が続く場合
                        iter.MoveNext()
                    Else
                        ' 不正な文字が出現した場合は終了
                        Exit While
                    End If
                End While

                ' 解析結果をTomlExpressionとして返す
                If EqualCharByte(iter.Current, &H27) Then ' '
                    iter.MoveNext()
                    Return New TomlExpression(TomlExpressionType.LiteralString, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' リテラル文字列の終了記号がない場合は例外をスロー
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("文字列の終了記号がありません。", source, iter)
                End If
            End If

            ' リテラル文字列の開始記号がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Charがリテラル文字列の文字であるかをチェックします。
        ''' リテラル文字列では、アポストロフィ（'）以外の文字が許可されています。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>リテラル文字列の文字ならば真。</returns>
        Private Function IsLiteralChar(u8c As U8Char?) As Boolean
            Return EqualCharByte(u8c, &H9) OrElse
                   RangeCharByte(u8c, &H20, &H26) OrElse
                   RangeCharByte(u8c, &H28, &H7E) OrElse
                   IsNonAscii(u8c)
        End Function

        '------------------------------
        ' 複数行リテラル文字列
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、複数行リテラル文字列を解析します。
        ''' 複数行リテラル文字列は3つのアポストロフィ（'''）で囲まれた文字列です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseMlLiteralString(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If MatchMlStringDelim(iter, &H27) Then
                ' 改行文字を読み捨てる
                MatchNewline(iter)

                ' 複数行リテラル文字列の内容を読み取る
                If Not MatchMlLiteralBody(iter) Then
                    Throw New TomlParseException("複数行リテラル文字列の内容がありません。", source, iter)
                End If

                ' 解析結果をTomlExpressionとして返す
                If MatchMlStringDelim(iter, &H27) Then
                    Return New TomlExpression(TomlExpressionType.MlLiteralString, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 複数行リテラル文字列の終了記号がない場合は例外をスロー
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("複数行リテラル文字列の終了記号がありません。", source, iter)
                End If
            End If

            ' 複数行リテラル文字列の開始記号がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 複数行リテラル文字列の内容を解析します。
        ''' ここでは、複数行リテラル文字列の本体部分を読み取ります。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行リテラル文字列の内容が正しく解析できた場合は真。</returns>
        Private Function MatchMlLiteralBody(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 複数行リテラル文字列の内容を読み取る
            While MatchMllContent(iter)
                ' 連続して複数行リテラル文字列の内容がマッチする限りループ
            End While

            While iter.HasNext()
                Dim enable = False
                Dim midIndex = iter.CurrentIndex
                If MatchMlQuotes(iter, &H27) Then
                    While MatchMllContent(iter)
                        enable = True
                    End While
                End If
                If Not enable Then
                    iter.SetCurrentIndex(midIndex)
                    Exit While
                End If
            End While

            ' 最後にアポストロフィが連続しているかをチェック
            Dim lastIndex = iter.CurrentIndex
            While iter.HasNext() AndAlso EqualCharByte(iter.Current, &H27)
                iter.MoveNext()
            End While
            Dim num = iter.CurrentIndex - lastIndex
            If num >= 3 AndAlso num <= 5 Then
                iter.SetCurrentIndex(lastIndex + (num - 3))
                Return True
            Else
                iter.SetCurrentIndex(lastIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、複数行リテラル文字列を解析します。
        ''' 複数行リテラル文字列は3つのアポストロフィ（'''）で囲まれた文字列です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行リテラル文字列内の文字ならば真。</returns>
        Private Function MatchMllContent(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            If IsMllChar(iter.Current) Then
                ' 複数行リテラル文字列内の文字がマッチした場合
                iter.MoveNext()
                Return True
            ElseIf MatchNewline(iter) Then
                ' 改行がマッチした場合
                Return True
            End If

            ' 複数行リテラル文字列内の文字がマッチしなかった場合は、イテレータの位置を元に戻す
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>複数行リテラル文字列の文字をチェックします。</summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>複数行リテラル文字列の文字の場合はTrue、それ以外はFalse。</returns>
        Private Function IsMllChar(u8c? As U8Char) As Boolean
            Return EqualCharByte(u8c, &H9) OrElse
                   RangeCharByte(u8c, &H20, &H26) OrElse
                   RangeCharByte(u8c, &H28, &H7E) OrElse
                   IsNonAscii(u8c)
        End Function

        '------------------------------
        ' 文字列共通
        '------------------------------
        ''' <summary>引数で指定されたU8Stringのイテレータから、1～2個、引用符が連続しているかをチェックします。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="quoByte">引用符のバイト値（アポストロフィまたはダブルコーテーション）。</param>
        ''' <returns>引用符が連続している場合は真。</returns>
        Private Function MatchMlQuotes(iter As U8String.U8StringIterator, quoByte As Byte) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 引用符が 1～2個連続しているかをチェック
            If EqualCharByte(iter.Current, quoByte) Then
                iter.MoveNext()
                If EqualCharByte(iter.Current, quoByte) Then
                    iter.MoveNext()
                End If
                Return True
            End If

            ' 引用符が連続していない場合はイテレータの位置を元に戻す
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、3つの引用符が連続しているかをチェックします。
        ''' 3つの引用符は複数行文字列の開始記号として使用されます。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="quoByte">引用符のバイト値（アポストロフィまたはダブルコーテーション）。</param>
        ''' <returns>3つの引用符が連続している場合は真。</returns>
        Private Function MatchMlStringDelim(iter As U8String.U8StringIterator, quoByte As Byte) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 3つの引用符が連続している場合は真を返す
            If EqualCharByte(iter.Current, quoByte) Then
                iter.MoveNext()
                If EqualCharByte(iter.Current, quoByte) Then
                    iter.MoveNext()
                    If EqualCharByte(iter.Current, quoByte) Then
                        iter.MoveNext()
                        Return True
                    End If
                End If
            End If

            ' 3つの引用符が連続していない
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

#End Region

#Region "数値"

        '------------------------------
        ' 整数
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、整数を解析します。
        ''' 整数は符号付き整数、16進数整数、8進数整数、2進数整数などを含みます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseInteger(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 16進数整数
            Dim hexExpr = ParseHexInt(source, iter)
            If hexExpr.Type <> TomlExpressionType.None Then
                Return hexExpr
            End If

            ' 8進数整数
            Dim octExpr = ParseOctInt(source, iter)
            If octExpr.Type <> TomlExpressionType.None Then
                Return octExpr
            End If

            ' 2進数整数
            Dim binExpr = ParseBinInt(source, iter)
            If binExpr.Type <> TomlExpressionType.None Then
                Return binExpr
            End If

            ' 符号付き整数
            Dim decExpr = ParseDecInt(source, iter)
            If decExpr.Type <> TomlExpressionType.None Then
                Return decExpr
            End If

            ' 整数でない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>n進数のプレフィックスをチェックします。</summary>
        ''' <param name="u8c1">最初のU8Char。</param>
        ''' <param name="u8c2">次のU8Char。</param>
        ''' <param name="byte1">1文字目。</param>
        ''' <param name="byte2">2文字目。</param>
        ''' <returns>n進数のプレフィックスの場合はTrue、それ以外はFalse。</returns>
        Private Function EqualNumPrefix(u8c1 As U8Char?, u8c2 As U8Char?, byte1 As Byte, byte2 As Byte) As Boolean
            If u8c1.HasValue AndAlso u8c2.HasValue Then
                Return (u8c1.Value.Raw0 = byte1 AndAlso u8c2.Value.Raw0 = byte2)
            End If
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、符号付き10進整数を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 符号付き整数は、符号（+または-）が先頭に来る形式です。
        ''' </remarks>
        Function ParseDecInt(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 符号が来た場合は読み進める
            ReadNumberSignIsPlus(iter)

            ' 符号付き整数として扱う
            If MatchUnsignedDecInt(iter) Then
                Return New TomlExpression(TomlExpressionType.DecInt, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                ' 符号があるが、整数が続かなかった場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、符号なし10進整数を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>有効な数値ならば真。</returns>
        ''' <remarks>
        ''' 符号なし整数は、0-9の数字が1つ以上続く形式です。
        ''' アンダースコア（_）は数字の区切りとして使用できます。
        ''' </remarks>
        Function MatchUnsignedDecInt(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex
            Dim firstChar = iter.Current

            If RangeCharByte(firstChar, &H31, &H39) Then ' 1-9の範囲
                ' 1から9の数字が最初に来る場合
                iter.MoveNext()
                Dim enable = False
                While iter.HasNext()
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        ' 0-9が続く場合
                        enable = True
                        iter.MoveNext()
                    ElseIf EqualCharByte(iter.Current, &H5F) AndAlso
                           RangeCharByte(iter.Peek(1), &H30, &H39) Then
                        ' アンダースコアが続き、その後に数字が来る場合
                        enable = True
                        iter.MoveNext()
                        iter.MoveNext()
                    Else
                        ' 不正な文字が出現した場合は終了
                        Exit While
                    End If
                End While

                ' 数字が1つ以上続いた場合は有効、そうでなければ無効
                If enable Then
                    Return True
                End If
            End If

            ' 数値1文字判定(0)
            iter.SetCurrentIndex(startIndex)
            If RangeCharByte(firstChar, &H30, &H39) Then ' 0-9の範囲
                iter.MoveNext()
                Return True
            End If

            ' 数字が来ない場合は終了
            Return False
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、16進数整数を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 16進数は0xで始まり、その後に0-9, A-Fが続く形式です。
        ''' </remarks>
        Public Function ParseHexInt(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If EqualNumPrefix(iter.Current, iter.Peek(1), &H30, &H78) Then ' 0x
                iter.MoveNext()
                iter.MoveNext()

                ' 16進数のプレフィックスがある場合、次に数字が続くかを確認
                If IsHexdig(iter.Current) Then
                    iter.MoveNext()
                    While iter.HasNext()
                        If IsHexdig(iter.Current) Then
                            ' 0-9, A-Fが続く場合
                            iter.MoveNext()
                        ElseIf EqualCharByte(iter.Current, &H5F) AndAlso IsHexdig(iter.Peek(1)) Then
                            ' アンダースコアが続き、その後に数字が来る場合
                            iter.MoveNext()
                            iter.MoveNext()
                        Else
                            ' 不正な文字が出現した場合は終了
                            Exit While
                        End If
                    End While

                    ' 16進数が1つ以上続いた場合はHexIntとして扱う
                    Return New TomlExpression(TomlExpressionType.HexInt, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 16進数のプレフィックスはあるが、数字が続かなかった場合
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("16進数のプレフィックスがあるが、数字が続きませんでした。", source, iter)
                End If
            End If

            ' 16進数のプレフィックスがない、または数字が続かなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、8進数整数を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 8進数は0oで始まり、その後に0-7が続く形式です。
        ''' </remarks>
        Public Function ParseOctInt(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If EqualNumPrefix(iter.Current, iter.Peek(1), &H30, &H6F) Then  ' 0o
                iter.MoveNext()
                iter.MoveNext()

                ' 8進数のプレフィックスがある場合、次に数字が続くかを確認
                If RangeCharByte(iter.Current, &H30, &H37) Then ' 0-7の範囲
                    iter.MoveNext()
                    While iter.HasNext()
                        If RangeCharByte(iter.Current, &H30, &H37) Then
                            ' 0-7が続く場合
                            iter.MoveNext()
                        ElseIf EqualCharByte(iter.Current, &H5F) AndAlso RangeCharByte(iter.Peek(1), &H30, &H37) Then
                            ' アンダースコアが続き、その後に数字が来る場合
                            iter.MoveNext()
                            iter.MoveNext()
                        Else
                            ' 不正な文字が出現した場合は終了
                            Exit While
                        End If
                    End While

                    ' 8進数が1つ以上続いた場合はOctIntとして扱う
                    Return New TomlExpression(TomlExpressionType.OctInt, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 8進数のプレフィックスはあるが、数字が続かなかった場合
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("8進数のプレフィックスがあるが、数字が続きませんでした。", source, iter)
                End If
            End If

            ' 8進数のプレフィックスがない、または数字が続かなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、2進数整数を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 8進数は0bで始まり、その後に0-1が続く形式です。
        ''' </remarks>
        Public Function ParseBinInt(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If EqualNumPrefix(iter.Current, iter.Peek(1), &H30, &H62) Then  ' 0b
                iter.MoveNext()
                iter.MoveNext()

                ' 2進数のプレフィックスがある場合、次に数字が続くかを確認
                If AnyCharByte(iter.Current, &H30, &H31) Then
                    iter.MoveNext()
                    While iter.HasNext()
                        If AnyCharByte(iter.Current, &H30, &H31) Then
                            ' 0-1が続く場合
                            iter.MoveNext()
                        ElseIf EqualCharByte(iter.Current, &H5F) AndAlso AnyCharByte(iter.Peek(1), &H30, &H31) Then
                            ' アンダースコアが続き、その後に数字が来る場合
                            iter.MoveNext()
                            iter.MoveNext()
                        Else
                            ' 不正な文字が出現した場合は終了
                            Exit While
                        End If
                    End While

                    ' 2進数が1つ以上続いた場合はBinIntとして扱う
                    Return New TomlExpression(TomlExpressionType.BinInt, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                Else
                    ' 2進数のプレフィックスはあるが、数字が続かなかった場合
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("2進数のプレフィックスがあるが、数字が続きませんでした。", source, iter)
                End If
            End If

            ' 2進数のプレフィックスがない、または数字が続かなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' 実数
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、浮動小数点数を解析します。
        ''' 浮動小数点数は、整数部と小数部、指数部を含む形式です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseFloat(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 特殊な浮動小数点数（InfやNaN）を解析
            Dim specialFloatExpr = ParseSpecialFloat(source, iter)
            If specialFloatExpr.Type <> TomlExpressionType.None Then
                Return specialFloatExpr
            End If

            ' 浮動小数点数の解析を試みる
            Dim decExpr = ParseDecInt(source, iter)
            If decExpr.Type <> TomlExpressionType.None Then
                Dim enable = False

                ' 浮動小数点数の指数部が続く場合
                If MatchExp(iter) Then
                    enable = True
                ElseIf iter.MoveNext()?.Raw0 = &H2E AndAlso MatchZeroPrefixableInt(iter) Then
                    MatchExp(iter)
                    enable = True
                End If

                ' 浮動小数点数が1つ以上続いた場合はFloatとして扱う
                If enable Then
                    Return New TomlExpression(TomlExpressionType.Float, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                End If
            End If

            ' 浮動小数点数の解析に失敗した場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、0始まりを許容する整数を解析します。
        ''' 0始まりの整数は、0から始まる数字で構成され、アンダースコアが含まれる場合があります。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>0始まりを許容する整数ならば真。</returns>
        ''' <remarks>
        ''' 例: 0123, 0_123, 0_1_2_3
        ''' </remarks>
        Function MatchZeroPrefixableInt(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex
            If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                iter.MoveNext()
                While iter.HasNext()
                    If RangeCharByte(iter.Current, &H30, &H39) Then
                        ' 0-9が続く場合
                        iter.MoveNext()
                    ElseIf EqualCharByte(iter.Current, &H5F) AndAlso
                           RangeCharByte(iter.Peek(1), &H30, &H39) Then
                        ' アンダースコアが続き、その後に数字が来る場合
                        iter.MoveNext()
                        iter.MoveNext()
                    Else
                        ' 不正な文字が出現した場合は終了
                        Exit While
                    End If
                End While
                Return True
            Else
                ' 数値が来ない場合は終了
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、浮動小数点数を解析します。
        ''' 浮動小数点数は、整数部と小数部、指数部を含む形式です。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Private Function MatchExp(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' e または E が来た場合は読み進める
            If AnyCharByte(iter.Current, &H65, &H45) Then
                iter.MoveNext()
            Else
                ' e または E が来なかった場合は終了
                Return False
            End If

            ' 浮動小数点数の指数部を解析する
            If MatchFloatExpPart(iter) Then
                Return True
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>引数で指定されたU8Stringから、浮動小数点数の指数部を解析します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>浮動小数点数の指数部ならば真。</returns>
        Private Function MatchFloatExpPart(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 符号が来た場合は読み進める
            ReadNumberSignIsPlus(iter)

            ' 浮動小数点数の指数部を解析する
            If MatchZeroPrefixableInt(iter) Then
                Return True
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、特殊な浮動小数点数（InfやNaN）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 特殊な浮動小数点数は、Inf（Infinity）やNaN（Not a Number）などを含みます。
        ''' </remarks>
        Function ParseSpecialFloat(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 符号が来た場合は読み進める
            ReadNumberSignIsPlus(iter)

            If MatchInf(iter) Then
                ' Inf（Infinity）を解析
                Return New TomlExpression(TomlExpressionType.Inf, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            ElseIf MatchNan(iter) Then
                ' NaN（Not a Number）を解析
                Return New TomlExpression(TomlExpressionType.Nan, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                ' InfやNaN以外の特殊な浮動小数点数ではない場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、Inf（Infinity）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Private Function MatchInf(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex
            If iter.MoveNext?.Raw0 = &H69 AndAlso
               iter.MoveNext?.Raw0 = &H6E AndAlso
               iter.MoveNext?.Raw0 = &H66 Then
                Return True
            Else
                ' Infが正しく解析できなかった場合は偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、NaN（Not a Number）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Private Function MatchNan(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex
            If iter.MoveNext?.Raw0 = &H6E AndAlso
               iter.MoveNext?.Raw0 = &H61 AndAlso
               iter.MoveNext?.Raw0 = &H6E Then
                Return True
            Else
                ' NANが正しく解析できなかった場合は偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

#End Region

#Region "真偽値"

        ''' <summary>
        ''' イテレーターの現在位置から、引数で指定されたバイト列と一致するかをチェックします。
        ''' 一致する場合はTrueを返し、イテレーターの位置を進めます。
        ''' 一致しない場合はFalseを返し、イテレーターの位置は変更されません。
        ''' </summary>
        ''' <param name="iter">イテレーター。</param>
        ''' <param name="bytes">比較するバイト配列。</param>
        ''' <returns>一致したら真。</returns>
        Private Function EqualBytes(iter As U8String.U8StringIterator, bytes() As Byte) As Boolean
            Dim startIndex = iter.CurrentIndex
            For Each b In bytes
                If iter.HasNext() AndAlso iter.Current.Value.Raw0 = b Then
                    iter.MoveNext()
                Else
                    ' 文字が一致しない場合は終了
                    iter.SetCurrentIndex(startIndex)
                    Return False
                End If
            Next
            ' 全ての文字が一致した場合はTrue
            Return True
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、真偽値（true または false）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 真偽値は、true または false のいずれかでなければなりません。
        ''' </remarks>
        Public Function ParseBoolean(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If EqualBytes(iter, {&H74, &H72, &H75, &H65}) Then ' true
                Return New TomlExpression(TomlExpressionType.BooleanLiteral, U8String.NewSlice(source, startIndex, 4))
            ElseIf EqualBytes(iter, {&H66, &H61, &H6C, &H73, &H65}) Then ' false
                Return New TomlExpression(TomlExpressionType.BooleanLiteral, U8String.NewSlice(source, startIndex, 5))
            Else
                ' true または false 以外の文字が来た場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If
        End Function

#End Region

#Region "日時"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、日時を解析します。
        ''' 日時はオフセット付き日時、ローカル日時、ローカル日付、ローカル時間のいずれかの形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 日時はISO 8601形式に準拠しています。
        ''' </remarks>
        Function ParseDateTime(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' オフセット付き日時を解析
            Dim offsetExpr = ParseOffsetDateTime(source, iter)
            If offsetExpr.Type <> TomlExpressionType.None Then
                Return offsetExpr
            End If

            ' ローカル日時を解析
            Dim localExpr = ParseLocalDateTime(source, iter)
            If localExpr.Type <> TomlExpressionType.None Then
                Return localExpr
            End If

            ' ローカル日付のみを解析
            Dim dateExpr = ParseLocalDate(source, iter)
            If dateExpr.Type <> TomlExpressionType.None Then
                Return dateExpr
            End If

            ' ローカル時間のみを解析
            Dim timeExpr = ParseLocalTime(source, iter)
            If timeExpr.Type <> TomlExpressionType.None Then
                Return timeExpr
            End If

            ' 日時の解析に失敗した場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、時間オフセットの部分を解析します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>時間オフセットの部分が解析できた場合は真。</returns>
        Private Function MatchTimeNumOffset(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 符号が来た場合は読み進める
            If AnyCharByte(iter.Current, &H2D, &H2B) Then ' - または +
                iter.MoveNext()
            Else
                ' 符号が来なかった場合は終了
                Return False
            End If

            ' 時間の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 時間の後にコロン（:）が来ない場合は終了
            If EqualCharByte(iter.Current, &H3A) Then ' :
                iter.MoveNext()
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 分の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            Return True
        End Function

        ''' <summary>引数で指定されたU8Stringから、時間オフセットの部分を解析します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>時間オフセットの部分が解析できた場合は真。</returns>
        Private Function MatchTimeOffset(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 時間オフセットの部分は、Z（Zulu time）または時間-分単位の形式で表されます。
            If EqualCharByte(iter.Current, &H5A) Then ' Z
                iter.MoveNext()
                Return True
            ElseIf MatchTimeNumOffset(iter) Then
                Return True
            End If

            ' 時間オフセットの解析に失敗した場合は、イテレータの位置を元に戻す
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        ''' <summary>引数で指定されたU8Stringから、時間-秒単位の小数部分を解析します。</summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>時間-秒単位の小数部分ならば真。</returns>
        ''' <remarks>
        ''' 時間-秒単位の小数部分は、ピリオド（.）に続く数字で構成されます。
        ''' </remarks>
        Function MatchTimeSecfrac(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            If iter.Current.HasValue AndAlso iter.Current.Value.Raw0 = &H2E Then ' .
                iter.MoveNext()

                Dim enable = False
                While iter.HasNext()
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        enable = True
                        iter.MoveNext()
                    Else
                        ' 不正な文字が出現した場合は終了
                        Exit While
                    End If
                End While

                ' 解析結果をTomlExpressionとして返す
                If enable Then
                    Return True
                Else
                    iter.SetCurrentIndex(startIndex)
                    Return False
                End If
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、部分的な時間（HH:MM:SS）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>部分的な時間が解析できた場合は真。</returns>
        ''' <remarks>
        ''' 部分的な時間は、HH:MM:SS[.SSS]の形式で表されます。
        ''' </remarks>
        Private Function MatchPartialTime(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 時間の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 時間の後にコロン（:）が来ない場合は終了
            If EqualCharByte(iter.Current, &H3A) Then ' :
                iter.MoveNext()
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 分の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 分の後にコロン（:）が来ない場合は終了
            If EqualCharByte(iter.Current, &H3A) Then ' :
                iter.MoveNext()
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 秒の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 秒の小数部分がある場合は読み進める
            MatchTimeSecfrac(iter)

            Return True
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、部分的な日付（YYYY-MM-DD）を解析します。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>部分的な日付が解析できた場合は真。</returns>
        ''' <remarks>
        ''' 部分的な日付は、YYYY-MM-DDの形式で表されます。
        ''' </remarks>
        Private Function MatchFullDate(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 年の部分を解析
            If Not MatchDigitTimes(iter, 4) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 年の後にコロン（-）が来ない場合は終了
            If EqualCharByte(iter.Current, &H2D) Then ' -
                iter.MoveNext()
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 月の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 月の後にコロン（-）が来ない場合は終了
            If EqualCharByte(iter.Current, &H2D) Then ' -
                iter.MoveNext()
            Else
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            ' 日の部分を解析
            If Not MatchDigitTimes(iter, 2) Then
                iter.SetCurrentIndex(startIndex)
                Return False
            End If

            Return True
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、完全な日時（YYYY-MM-DDTHH:MM:SS[.SSS][Z|±HH:MM]）を解析します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>完全な日時が解析できた場合は真。</returns>
        ''' <remarks>
        ''' 完全な日時は、ISO 8601形式の日付と時間を表します。
        ''' </remarks>
        Private Function MatchFullTime(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 時間の部分が解析できた場合は、時間オフセットを解析して返す
            If MatchPartialTime(iter) AndAlso MatchTimeOffset(iter) Then
                Return True
            End If

            ' 時間オフセットがない場合は、イテレーターの位置を元に戻す
            iter.SetCurrentIndex(startIndex)
            Return False
        End Function

        '------------------------------
        ' オフセット日時
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、オフセット日時を解析します。
        ''' オフセット日時は、YYYY-MM-DDTHH:MM:SS[.SSS][Z|±HH:MM]の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' オフセット日時は、ISO 8601形式の日付と時間を表し、タイムゾーンオフセットを含みます。
        ''' </remarks>
        Function ParseOffsetDateTime(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If MatchFullDate(iter) Then
                If AnyCharByte(iter.Current, &H54, &H20) Then ' Tまたは空白
                    iter.MoveNext() ' Tまたは空白を読み飛ばす
                    If MatchFullTime(iter) Then
                        Return New TomlExpression(TomlExpressionType.OffsetDateTime, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                    End If
                End If
            End If

            ' オフセット日時の解析ができなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' ローカル日時
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、ローカル日時を解析します。
        ''' ローカル日時は、YYYY-MM-DDTHH:MM:SS[.SSS]の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' ローカル日時は、ISO 8601形式の日付と時間を表します。
        ''' </remarks>
        Function ParseLocalDateTime(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If MatchFullDate(iter) Then
                If AnyCharByte(iter.Current, &H54, &H20) Then ' Tまたは空白
                    iter.MoveNext() ' Tまたは空白を読み飛ばす
                    If MatchPartialTime(iter) Then
                        Return New TomlExpression(TomlExpressionType.LocalDateTime, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
                    End If
                End If
            End If

            ' 日時の解析ができなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' ローカル日付
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、ローカル日付を解析します。
        ''' ローカル日付は、YYYY-MM-DDの形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' ローカル日付は、YYYY-MM-DDの形式で表されます。
        ''' </remarks>
        Function ParseLocalDate(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If MatchFullDate(iter) Then
                Return New TomlExpression(TomlExpressionType.LocalDate, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            End If

            ' 日付の解析ができなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' ローカル時間
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、ローカル時間を解析します。
        ''' ローカル時間は、時:分:秒の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' ローカル時間は、HH:MM:SS[.SSS]の形式で表されます。
        ''' </remarks>
        Function ParseLocalTime(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If MatchPartialTime(iter) Then
                Return New TomlExpression(TomlExpressionType.LocalTime, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            End If

            ' 時間の解析ができなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

#End Region

#Region "配列"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、配列を解析します。
        ''' 配列は、値のリストであり、カンマ（,）で区切られます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 配列は、[ と ] で囲まれた値のリストです。
        ''' </remarks>
        Function ParseArray(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            If EqualCharByte(iter.Current, &H5B) Then
                ' 配列の始端（[）が来た場合は読み進める
                iter.MoveNext()

                ' 配列の値を解析
                Dim expressions As New List(Of TomlExpression)()
                Dim arrayValues = ParseArrayValues(source, iter, expressions)

                ' 空白文字、コメント、改行をスキップ)
                ParseWsCommentNewline(source, iter)

                ' 配列の終端（]）が来た場合は読み進める
                If EqualCharByte(iter.Current, &H5D) Then
                    iter.MoveNext()
                    Return New TomlExpression(TomlExpressionType.Array, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex), expressions.ToArray())
                Else
                    ' 配列の終端が来なかった場合は例外を投げる
                    Throw New TomlParseException("配列の終端（]）が見つかりませんでした。", source, iter)
                End If
            End If

            ' 配列の始端が来なかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、配列の値を解析します。
        ''' 配列は、値のリストであり、カンマ（,）で区切られます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseArrayValues(source As U8String, iter As U8String.U8StringIterator, expressions As List(Of TomlExpression)) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 空白文字、コメント、改行をスキップ)
            ParseWsCommentNewline(source, iter)

            ' 配列の値を解析
            Dim valueExpr = ParseVal(source, iter)
            If valueExpr.Type <> TomlExpressionType.None Then
                expressions.Add(valueExpr)
            Else
                ' 値が解析できなかった場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If

            ' 空白文字、コメント、改行をスキップ)
            ParseWsCommentNewline(source, iter)

            ' カンマが来た場合は次の値を解析
            If EqualCharByte(iter.Current, &H2C) Then
                iter.MoveNext()

                ' 次の値を再帰的に解析
                ParseArrayValues(source, iter, expressions)
            End If

            ' 配列の値をすべて解析した後、配列の値の式を作成
            Return New TomlExpression(TomlExpressionType.ArrayValues, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
        End Function

        ''' <summary>
        ''' 空白文字、コメント、改行を解析します。
        ''' これらはTomlExpressionの一部ではなく、単に解析のために使用されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>空白文字、コメント、改行があれば真。</returns>
        ''' <remarks>
        ''' 空白文字はスペースやタブ、改行はLFやCRLFを含みます。
        ''' コメントは#で始まる行で、行末までがコメントとして扱われます。
        ''' </remarks>
        Function ParseWsCommentNewline(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            While iter.HasNext()
                Dim midIndex = iter.CurrentIndex
                If AnyCharByte(iter.Current, &H20, &H9) Then
                    ' 空白文字を読み捨て
                    iter.MoveNext()
                    Continue While
                Else
                    ' コメント、改行を読み捨て
                    ParseComment(source, iter)
                    If MatchNewline(iter) Then
                        Continue While
                    End If
                End If

                iter.SetCurrentIndex(midIndex)
                Exit While
            End While

            If iter.CurrentIndex > startIndex Then
                ' 空白文字、コメント、改行があった場合はTomlExpressionを返す
                Return New TomlExpression(TomlExpressionType.WsCommentNewline, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
            Else
                ' 空白文字、コメント、改行がなかった場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If
        End Function

#End Region

#Region "テーブル"

        ''' <summary>
        ''' 引数で指定されたU8Stringから、テーブルを解析します。
        ''' テーブルは、標準テーブルとインラインテーブルの2種類があります。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 標準テーブルは [table] の形式で表され、インラインテーブルは {key = value, key2 = value2} の形式で表されます。
        ''' </remarks>
        Function ParseTable(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 配列テーブルを解析
            Dim arrayTable = ParseArrayTable(source, iter)
            If arrayTable.Type <> TomlExpressionType.None Then
                Return arrayTable
            End If

            ' 標準テーブルを解析
            Dim stdTable = ParseStdTable(source, iter)
            If stdTable.Type <> TomlExpressionType.None Then
                Return stdTable
            End If

            ' どちらのテーブルも解析できなかった場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' 標準テーブル
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、標準テーブルを解析します。
        ''' 標準テーブルは、[table] の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 標準テーブルは、[と]で囲まれた名前空間を持つテーブルです。
        ''' </remarks>
        Private Function ParseStdTable(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' 標準テーブルの開始を確認 [
            If MatchOneCharAndLastSkipWs(iter, &H5B) Then
                ' キーを解析
                Dim key = ParseKey(source, iter)
                If key.Type = TomlExpressionType.None Then
                    ' キーが解析できなかった場合は空の式を返す
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("テーブルのキーが正しく解析できませんでした。", source, iter)
                End If

                ' テーブルの終了を確認 ]
                If MatchOneCharAndPrevSkipWs(iter, &H5D) Then
                    Return New TomlExpression(TomlExpressionType.Table, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex), New TomlExpression() {key})
                Else
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("テーブルの終了が正しく解析できませんでした。", source, iter)
                End If
            End If

            ' テーブルの開始がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        '------------------------------
        ' インラインテーブル
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、インラインテーブルを解析します。
        ''' インラインテーブルは、{key = value, key2 = value2} の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' インラインテーブルは、{と}で囲まれた名前と値のペアを持つテーブルです。
        ''' </remarks>
        Function ParseInlineTable(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' インラインテーブルの開始を確認 {
            If MatchOneCharAndLastSkipWs(iter, &H7B) Then
                ' キー、値を解析
                Dim expressions As New List(Of TomlExpression)()
                ParseInlineTableKeyVals(source, iter, expressions)

                ' インラインテーブルの終了を確認 }
                If MatchOneCharAndPrevSkipWs(iter, &H7D) Then
                    Return New TomlExpression(TomlExpressionType.InlineTable, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex), expressions.ToArray())
                Else
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("インラインテーブルの終了が正しく解析できませんでした。", source, iter)
                End If
            End If

            ' インラインテーブルの開始がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、インラインテーブルのキーと値の区切り文字（,）を解析します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>区切り文字が見つかった場合は真、それ以外は偽。</returns>
        Private Function MatchlineTableSep(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 空白文字を読み捨て
            MatchWs(iter)

            ' 区切り文字（,）をチェック
            Dim res = EqualCharByte(iter.Current, &H2C)
            iter.MoveNext()

            ' 空白文字を読み捨て
            MatchWs(iter)

            If res Then
                ' 区切り文字が見つかった場合は真を返す
                Return True
            Else
                ' 区切り文字が見つからなかった場合はイテレータの位置を戻す
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringから、インラインテーブルのキーと値のペアを解析します。
        ''' インラインテーブルは、{key = value, key2 = value2} の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="expressions">解析結果の式を格納するリスト。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        Function ParseInlineTableKeyVals(source As U8String, iter As U8String.U8StringIterator, expressions As List(Of TomlExpression)) As TomlExpression
            Dim startIndex = iter.CurrentIndex

            ' キーを解析
            Dim keyVal = ParseKeyval(source, iter)
            If keyVal.Type <> TomlExpressionType.None Then
                expressions.Add(keyVal)
            Else
                ' キーが解析できなかった場合は空の式を返す
                iter.SetCurrentIndex(startIndex)
                Return TomlExpression.Empty
            End If

            ' 区切り文字か判定
            If MatchlineTableSep(iter) Then
                Dim nextKeyVal = ParseInlineTableKeyVals(source, iter, expressions)
                If nextKeyVal.Type = TomlExpressionType.None Then
                    iter.SetCurrentIndex(startIndex)
                    Return TomlExpression.Empty
                End If
            End If

            ' 解析結果をTomlExpressionとして返す
            Return New TomlExpression(TomlExpressionType.InlineTableKeyvals, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex))
        End Function

        '------------------------------
        ' 配列テーブル
        '------------------------------
        ''' <summary>
        ''' 引数で指定されたU8Stringから、配列テーブルを解析します。
        ''' 配列テーブルは、[[配列名]]の形式で表されます。
        ''' </summary>
        ''' <param name="source">解析対象のU8String。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>解析結果のTomlExpression。</returns>
        ''' <remarks>
        ''' 配列テーブルは、[[と]]で囲まれた名前を持つテーブルです。
        ''' </remarks>
        Function ParseArrayTable(source As U8String, iter As U8String.U8StringIterator) As TomlExpression
            Dim startIndex = iter.CurrentIndex
            If MatchArrayTableOpen(iter) Then
                ' キーを解析
                Dim key = ParseKey(source, iter)
                If key.Type = TomlExpressionType.None Then
                    ' キーが解析できなかった場合は空の式を返す
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("配列テーブルのキーが正しく解析できませんでした。", source, iter)
                End If

                ' 配列テーブルの終了を確認
                If MatchArrayTableClose(iter) Then
                    Return New TomlExpression(TomlExpressionType.ArrayTable, U8String.NewSlice(source, startIndex, iter.CurrentIndex - startIndex), key.Contents)
                Else
                    iter.SetCurrentIndex(startIndex)
                    Throw New TomlParseException("配列テーブルの終了が正しく解析できませんでした。", source, iter)
                End If
            End If

            ' 配列テーブルの開始がない場合は空の式を返す
            iter.SetCurrentIndex(startIndex)
            Return TomlExpression.Empty
        End Function

        ''' <summary>引数で指定されたU8Stringから、配列テーブルの開始を解析します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>配列テーブルの開始ならば真。</returns>
        Private Function MatchArrayTableOpen(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex
            If iter.MoveNext?.Raw0 = &H5B AndAlso iter.MoveNext?.Raw0 = &H5B Then
                ' 空白文字を読み捨て
                MatchWs(iter)

                Return True
            Else
                ' 配列テーブルの開始が正しく解析できなかった場合は偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>引数で指定されたU8Stringから、配列テーブルの終了を解析します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>配列テーブルの終了ならば真。</returns>
        Private Function MatchArrayTableClose(iter As U8String.U8StringIterator) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 空白文字を読み捨て
            MatchWs(iter)

            If iter.MoveNext?.Raw0 = &H5D AndAlso iter.MoveNext?.Raw0 = &H5D Then
                Return True
            Else
                ' 配列テーブルの終了が正しく解析できなかった場合は偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

#End Region

#Region "標準要素と共通処理"

        ''' <summary>
        ''' 引数のU8Charが指定されたバイト値と一致するかどうかをチェックします。
        ''' 一致する場合は真を返します。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <param name="charByte">対象のバイト値。</param>
        ''' <returns>対象のバイト値と一致する場合は真。</returns>
        Private Function EqualCharByte(u8c As U8Char?, charByte As Byte) As Boolean
            If u8c.HasValue Then
                Return (u8c.Value.Raw0 = charByte)
            End If
            Return False
        End Function

        ''' <summary>
        ''' 引数のU8Charが指定されたバイト範囲内にあるかどうかをチェックします。
        ''' 範囲はlowByteからhiByteまでです。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <param name="lowByte">範囲の下限バイト値。</param>
        ''' <param name="hiByte">範囲の上限バイト値。</param>
        ''' <returns>範囲内にある場合は真、それ以外は偽。</returns>
        Private Function RangeCharByte(u8c As U8Char?, lowByte As Byte, hiByte As Byte) As Boolean
            If u8c.HasValue Then
                Return u8c.Value.Raw0 >= lowByte AndAlso u8c.Value.Raw0 <= hiByte
            End If
            Return False
        End Function

        ''' <summary>
        ''' 引数のU8Charが指定されたバイト値のいずれかと一致するかどうかをチェックします。
        ''' 一致する場合は真を返します。
        ''' </summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <param name="tarByte">対象のバイト値の配列。</param>
        ''' <returns>対象のバイト値のいずれかと一致する場合は真、それ以外は偽。</returns>
        Private Function AnyCharByte(u8c As U8Char?, ParamArray tarByte() As Byte) As Boolean
            If u8c.HasValue Then
                For b As Integer = 0 To tarByte.Length - 1
                    If u8c.Value.Raw0 = tarByte(b) Then
                        Return True
                    End If
                Next
            End If
            Return False
        End Function

        ''' <summary>アルファベット文字（a-z, A-Z）をチェックします。</summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>アルファベット文字の場合はTrue、それ以外はFalse。</returns>
        Private Function IsAlpha(u8c As U8Char?) As Boolean
            ' アルファベット文字（a-z, A-Z）をチェック
            Return RangeCharByte(u8c, &H41, &H5A) OrElse RangeCharByte(u8c, &H61, &H7A)
        End Function

        ''' <summary>16進数文字（0-9, A-F, a-f）をチェックします。</summary>
        ''' <param name="u8c">U8Char。</param>
        ''' <returns>16進数文字の場合はTrue、それ以外はFalse。</returns>
        Private Function IsHexdig(u8c As U8Char?) As Boolean
            ' 16進数文字（0-9, A-F）をチェック
            Return RangeCharByte(u8c, &H30, &H39) OrElse ' 0-9の範囲
                   RangeCharByte(u8c, &H41, &H46) OrElse ' A-Fの範囲
                   RangeCharByte(u8c, &H61, &H66) ' a-fの範囲
        End Function

        ''' <summary>指定回数の 数字文字かチェックします。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="count">文字数</param>
        ''' <returns>指定回数の数字文字の場合はTrue、それ以外はFalse。</returns>
        Private Function MatchDigitTimes(iter As U8String.U8StringIterator, count As Integer) As Boolean
            For i As Integer = 0 To count - 1
                If Not RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                    Return False
                End If
                iter.MoveNext()
            Next
            Return True
        End Function

        ''' <summary>指定回数の 16進数文字かチェックします。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="count">文字数</param>
        ''' <returns>指定回数の 16進数文字の場合はTrue、それ以外はFalse。</returns>
        Private Function MatchHexdigTimes(iter As U8String.U8StringIterator, count As Integer) As Boolean
            For i As Integer = 0 To count - 1
                If Not IsHexdig(iter.Current) Then
                    Return False
                End If
                iter.MoveNext()
            Next
            Return True
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから空白を読み捨ててから
        ''' 対象文字かどうかをチェックし、対象文字ならば真を返します。
        ''' 対象文字でなければイテレータの位置を戻して偽を返します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="charByte">対象文字。</param>
        ''' <returns>対象文字ならば真。</returns>
        Private Function MatchOneCharAndPrevSkipWs(iter As U8String.U8StringIterator, charByte As Byte) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 空白文字を読み捨て
            MatchWs(iter)

            ' 対象文字なら真
            If iter.MoveNext?.Raw0 = charByte Then
                Return True
            Else
                ' 対象文字が来なかった場合はイテレータの位置を戻して偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから対象文字かどうかをチェックし、
        ''' 対象文字ならば空白文字を読み捨てて真を返します。
        ''' 対象文字でなければイテレータの位置を戻して偽を返します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="charByte">対象文字。</param>
        ''' <returns>対象文字ならば真。</returns>
        Private Function MatchOneCharAndLastSkipWs(iter As U8String.U8StringIterator, charByte As Byte) As Boolean
            Dim startIndex = iter.CurrentIndex

            ' 対象文字なら空白文字を読み捨てて真
            If iter.MoveNext?.Raw0 = charByte Then
                ' 空白文字を読み捨て
                MatchWs(iter)
                Return True
            Else
                ' 対象文字が来なかった場合はイテレータの位置を戻して偽
                iter.SetCurrentIndex(startIndex)
                Return False
            End If
        End Function

        ''' <summary>
        ''' 16進数文字を数値に変換します。
        ''' 引数のバイトは、ASCIIコードの16進数文字（0-9, A-F, a-f）である必要があります。
        ''' それ以外の文字が来た場合は例外をスローします。
        ''' </summary>
        ''' <param name="srcByte">変換する文字。</param>
        ''' <returns>変換した値。</returns>
        Private Function ToByteHex(srcByte As Byte) As Byte
            If srcByte >= &H30 AndAlso srcByte <= &H39 Then
                Return CByte(srcByte - &H30) ' 0-9
            ElseIf srcByte >= &H41 AndAlso srcByte <= &H46 Then
                Return CByte(srcByte - &H41 + 10) ' A-F
            ElseIf srcByte >= &H61 AndAlso srcByte <= &H66 Then
                Return CByte(srcByte - &H61 + 10) ' a-f
            End If
            Throw New TomlParseException("無効な16進数文字です。")
        End Function

        ''' <summary>U8StringのイテレータからUTF-8の文字を取得し、バイト領域に追加します。</summary>
        ''' <param name="buf">バイト列を格納するリスト。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="times">追加する16進数の文字数。</param>
        Private Sub AddHexIter(buf As List(Of Byte), iter As U8String.U8StringIterator, times As Integer)
            Dim b As Byte = 0
            For i As Integer = 0 To times - 1
                b = (b << 4) Or ToByteHex(iter.Current.Value.Raw0)
                iter.MoveNext()
                If (i And 1) <> 0 Then
                    If b <> 0 Then
                        buf.Add(b)
                    End If
                    b = 0
                End If
            Next
        End Sub

        ''' <summary>
        ''' バイトバッファにU8Charを追加します。
        ''' U8CharはUTF-8の文字を表し、1〜4バイトで構成されます。
        ''' </summary>
        ''' <param name="buf">バイトバッファ。</param>
        ''' <param name="u8c">追加する文字。</param>
        <Extension()>
        Private Sub AppendU8Char(buf As List(Of Byte), u8c As U8Char)
            ' U8Charをバイト列に追加
            Select Case u8c.Size
                Case 1
                    buf.Add(u8c.Raw0)
                Case 2
                    buf.Add(u8c.Raw0)
                    buf.Add(u8c.Raw1)
                Case 3
                    buf.Add(u8c.Raw0)
                    buf.Add(u8c.Raw1)
                    buf.Add(u8c.Raw2)
                Case 4
                    buf.Add(u8c.Raw0)
                    buf.Add(u8c.Raw1)
                    buf.Add(u8c.Raw2)
                    buf.Add(u8c.Raw3)
            End Select
        End Sub

        ''' <summary>U8StringのイテレータからUTF-8の文字を取得し、バイト領域に追加します。</summary>
        ''' <param name="buf">バイト列を格納するリスト。</param>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        <Extension()>
        Private Sub AppendEscapeBytes(buf As List(Of Byte), iter As U8String.U8StringIterator)
            Dim nc = iter.MoveNext()
            Select Case nc.Value.Raw0
                Case &H22, &H5C
                    buf.Add(nc.Value.Raw0)
                Case &H62
                    buf.Add(&H8)
                Case &H66
                    buf.Add(&HC)
                Case &H6E
                    buf.Add(&HA)
                Case &H72
                    buf.Add(&HD)
                Case &H74
                    buf.Add(&H9)
                Case &H75
                    AddHexIter(buf, iter, 4)
                Case &H55
                    AddHexIter(buf, iter, 8)
                Case Else
                    Throw New TomlParseException("無効なエスケープシーケンスです。")
            End Select
        End Sub

        ''' <summary>引数で指定されたU8Stringのイテレータから、数値の符号がプラスかどうかを読み取ります。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>符号がプラスの場合は真、それ以外は偽。</returns>
        Private Function ReadNumberSignIsPlus(iter As U8String.U8StringIterator) As Boolean
            ' 符号がプラス（+）かマイナス（-）かをチェック
            If iter.Current.HasValue Then
                If iter.Current.Value.Raw0 = &H2D Then
                    iter.MoveNext()
                    Return False
                ElseIf iter.Current.Value.Raw0 = &H2B Then
                    iter.MoveNext()
                    Return True
                End If
            End If

            ' +、-以外は正の数値とする
            Return True
        End Function

#End Region

    End Module

End Namespace
