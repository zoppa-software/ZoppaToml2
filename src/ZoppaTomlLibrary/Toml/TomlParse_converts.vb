Option Strict On
Option Explicit On
Imports ZoppaTomlLibrary.Strings

Namespace Toml

    Partial Module TomlParser

#Region "文字列コンバート"

        ''' <summary>
        ''' U8Stringから基本文字列を解析します。
        ''' 基本文字列はダブルコーテーション（"）で囲まれた文字列で、エスケープシーケンスを含むことができます。
        ''' この関数は、基本文字列の開始記号がある場合にのみ動作します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>文字列。</returns>
        Function ConvertToBasicString(iter As U8String.U8StringIterator) As U8String
            Dim startIndex = iter.CurrentIndex
            If EqualCharByte(iter.Current, &H22) Then ' "
                Dim buf As New List(Of Byte)()
                iter.MoveNext()

                While iter.HasNext()
                    Dim c = iter.MoveNext().Value

                    ' 基本文字列の文字を解析してバイトバッファに追加
                    ' 1. エスケープシーケンスの文字ならば解析して追加
                    ' 2. 非エスケープの基本文字ならば追加
                    ' 3. 基本文字列の終了記号（"）ならば終了
                    ' 4. 基本文字列の終了記号でない場合は例外をスロー
                    If IsEscapedSeqChar(c, iter.Current) Then
                        buf.AppendEscapeBytes(iter) ' 1
                    ElseIf Not EqualCharByte(c, &H22) Then　' " の場合
                        buf.AppendU8Char(c)         ' 2
                    ElseIf EqualCharByte(c, &H22) Then ' "
                        Return U8String.NewStringChangeOwner(buf.ToArray()) ' 3
                    Else
                        iter.SetCurrentIndex(startIndex)    ' 4
                        Throw New TomlParseException("基本文字列の終了記号がありません。")
                    End If
                End While
            End If
            iter.SetCurrentIndex(startIndex)
            Throw New TomlParseException("基本文字列の開始記号がありません。")
        End Function

        ''' <summary>
        ''' 複数行基本文字列の内容をU8Stringに変換します。
        ''' ここでは、複数行基本文字列の内容を読み取り、U8Stringとして返します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行基本文字列の内容を含むU8String。</returns>
        Function ConvertToMlBasicString(iter As U8String.U8StringIterator) As U8String
            If MatchMlStringDelim(iter, &H22) Then
                Dim buf As New List(Of Byte)()

                ' 改行文字を読み捨てる
                MatchNewline(iter)

                While iter.HasNext()
                    Dim c = iter.Current

                    If IsEscapedSeqChar(c, iter.Peek(1)) Then
                        ' エスケープシーケンスを解析
                        iter.MoveNext()
                        buf.AppendEscapeBytes(iter)
                    ElseIf AnyCharByte(c, &H20, &H9, &H21) OrElse
                           RangeCharByte(c.Value, &H23, &H5B) OrElse
                           RangeCharByte(c.Value, &H5D, &H7E) Then
                        ' エスケープされていない文字
                        iter.MoveNext()
                        buf.Add(c.Value.Raw0)
                    ElseIf c.Value.Raw0 = &HA Then
                        ' LF (0x0A) の場合
                        iter.MoveNext()
                        buf.Add(c.Value.Raw0)
                    ElseIf c.Value.Raw0 = &HD AndAlso iter.Peek(1).Value.Raw0 = &HA Then
                        ' CR (0x0D) 、LF (0x0A) の場合
                        buf.Add(iter.Current.Value.Raw0)
                        iter.MoveNext()
                        buf.Add(iter.Current.Value.Raw0)
                        iter.MoveNext()
                    ElseIf IsNonAscii(c) Then
                        ' 非ASCII文字はそのまま追加
                        iter.MoveNext()
                        buf.AppendU8Char(c.Value)
                    End If

                    ' 行末改行は読み捨て
                    MatchMlbEscapedNl(iter)

                    ' 複数行基本文字列の終了記号が来た場合は終了
                    If MatchMlStringDelim(iter, &H22) Then
                        While iter.HasNext() AndAlso EqualCharByte(iter.Current, &H22)
                            buf.Add(&H22)
                            iter.MoveNext()
                        End While
                        Return U8String.NewStringChangeOwner(buf.ToArray())
                    End If

                    ' ダブルコーテーションが連続している場合は、ダブルコーテーションを追加
                    If MatchMlQuotes(iter, &H22) Then
                        buf.Add(&H22)
                    End If
                End While
            End If
            Throw New TomlParseException("複数行基本文字列の開始記号がありません。")
        End Function

        ''' <summary>
        ''' U8Stringのイテレータからリテラル文字列を取得します。
        ''' リテラル文字列はアポストロフィ（'）で囲まれた文字列で、エスケープシーケンスを含まない文字列です。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>文字列。</returns>
        Function ConvertToLiteralString(iter As U8String.U8StringIterator) As U8String
            Dim startIndex = iter.CurrentIndex
            If EqualCharByte(iter.Current, &H27) Then ' '
                Dim buf As New List(Of Byte)()
                iter.MoveNext()

                While iter.HasNext()
                    Dim c = iter.MoveNext().Value

                    ' リテラル文字列の文字を解析してバイトバッファに追加
                    ' 1. 非エスケープの基本文字ならば追加
                    ' 2. リテラル文字列の終了記号（'）ならば終了
                    ' 3. リテラル文字列の終了記号でない場合は例外をスロー
                    If IsLiteralChar(c) Then
                        buf.AppendU8Char(c) ' 1
                    ElseIf EqualCharByte(c, &H27) Then ' '
                        Return U8String.NewStringChangeOwner(buf.ToArray()) ' 2
                    Else
                        iter.SetCurrentIndex(startIndex)    ' 3
                        Throw New TomlParseException("リテラル文字列の終了記号がありません。")
                    End If
                End While
            End If
            iter.SetCurrentIndex(startIndex)
            Throw New TomlParseException("リテラル文字列の開始記号がありません。")
        End Function

        ''' <summary>
        ''' 複数行リテラル文字列の内容をU8Stringに変換します。
        ''' ここでは、複数行リテラル文字列の内容を読み取り、U8Stringとして返します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>複数行リテラル文字列の内容を含むU8String。</returns>
        Function ConvertToMlLiteralString(iter As U8String.U8StringIterator) As U8String
            If MatchMlStringDelim(iter, &H27) Then
                Dim buf As New List(Of Byte)()

                ' 改行文字を読み捨てる
                MatchNewline(iter)

                While iter.HasNext()
                    Dim c = iter.Current

                    If EqualCharByte(c.Value, &H9) OrElse
                       RangeCharByte(c.Value, &H20, &H26) OrElse
                       RangeCharByte(c.Value, &H28, &H7E) Then
                        ' 文字
                        iter.MoveNext()
                        buf.Add(c.Value.Raw0)
                    ElseIf c.Value.Raw0 = &HA Then
                        ' LF (0x0A) の場合
                        iter.MoveNext()
                        buf.Add(c.Value.Raw0)
                    ElseIf c.Value.Raw0 = &HD AndAlso iter.Peek(1).Value.Raw0 = &HA Then
                        ' CR (0x0D) 、LF (0x0A) の場合
                        buf.Add(iter.Current.Value.Raw0)
                        iter.MoveNext()
                        buf.Add(iter.Current.Value.Raw0)
                        iter.MoveNext()
                    ElseIf IsNonAscii(c) Then
                        ' 非ASCII文字はそのまま追加
                        iter.MoveNext()
                        buf.AppendU8Char(c.Value)
                    End If

                    ' 複数行基本文字列の終了記号が来た場合は終了
                    If MatchMlStringDelim(iter, &H27) Then
                        While iter.HasNext() AndAlso EqualCharByte(iter.Current, &H27)
                            buf.Add(&H27)
                            iter.MoveNext()
                        End While
                        Return U8String.NewStringChangeOwner(buf.ToArray())
                    End If

                    ' アポストロフィが連続している場合は、アポストロフィを追加
                    If MatchMlQuotes(iter, &H27) Then
                        buf.Add(&H27)
                    End If
                End While
            End If
            Throw New TomlParseException("複数行リテラル文字列の開始記号がありません。")
        End Function

#End Region

#Region "整数コンバート"

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、符号付き整数を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換された符号付き整数。</returns>
        ''' <remarks>
        ''' 符号がある場合は、符号に応じて数値を計算します。
        ''' </remarks>
        Function ConvertToDecInt(iter As U8String.U8StringIterator) As Long
            ' 符号を読み進める
            Dim isPlus = ReadNumberSignIsPlus(iter)

            ' 文字を数値化
            Dim res As ULong = 0
            While iter.HasNext()
                With iter.Current
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        res = res * 10UL + CULng(iter.Current.Value.Raw0 - &H30)
                    End If
                End With
                iter.MoveNext()
            End While

            ' 符号を考慮して数値を計算
            If isPlus Then
                If res < &H8000000000000000UL Then
                    Return CLng(res)
                Else
                    Throw New OverflowException("符号付き整数の範囲を超えました。")
                End If
            Else
                If res <= &H8000000000000000UL Then
                    Return CLng(-res)
                Else
                    Throw New OverflowException("符号付き整数の範囲を超えました。")
                End If
            End If
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、符号付き整数を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="numbase">数値の基数（10進数、16進数など）。</param>
        ''' <returns>変換された符号付き整数。</returns>
        ''' <remarks>
        ''' 符号がある場合は、符号に応じて数値を計算します。
        ''' </remarks>
        Function ConvertToDecInt(iter As U8String.U8StringIterator, numbase As Integer) As ULong
            ' 接頭辞を読み捨て
            iter.MoveNext()
            iter.MoveNext()

            ' 文字を数値化
            Dim res As ULong = 0
            While iter.HasNext()
                With iter.Current
                    If IsHexdig(iter.Current) Then
                        res = (res << numbase) + ToByteHex(iter.Current.Value.Raw0)
                    End If
                End With
                iter.MoveNext()
            End While

            ' 符号を考慮して数値を計算
            Return res
        End Function

#End Region

#Region "実数コンバート"

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、浮動小数点数を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換された浮動小数点数。</returns>
        ''' <remarks>
        ''' 符号がある場合は、符号に応じて数値を計算します。
        ''' </remarks>
        Function ConvertToFloat(iter As U8String.U8StringIterator) As Double
            ' 符号を読み進める
            Dim isPlus = ReadNumberSignIsPlus(iter)

            ' 文字を数値化
            Dim res As ULong = 0
            Dim decfig As Integer = -1
            While iter.HasNext()
                With iter.Current
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        res = res * 10UL + CULng(iter.Current.Value.Raw0 - &H30)
                        ' 小数点以下の桁数をカウント
                        If decfig >= 0 Then
                            decfig += 1
                        End If
                    ElseIf iter.Current.Value.Raw0 = &H2E Then
                        decfig = 0
                    ElseIf iter.Current.Value.Raw0 <> &H5F Then
                        ' 小数点以外の数字が来た場合は終了
                        Exit While
                    End If
                End With

                iter.MoveNext()
            End While

            ' 小数点がなかった場合は、整数部の桁数を0に設定
            If decfig < 0 Then
                decfig = 0
            End If

            ' 指数部がある場合は、10のべき乗を調整
            ' 例えば、1.23e4は1.23 * 10^4を意味します。
            ' 1. 指数部の符号をチェック
            ' 2. 文字を数値化
            ' 3. 先の指数部の値を調整
            If iter.HasNext() AndAlso AnyCharByte(iter.Current, &H65, &H45) Then
                iter.MoveNext()

                Dim expSign = ReadNumberSignIsPlus(iter)    ' 1
                Dim expValue = ToFloatDecInt(iter)          ' 2
                decfig += If(expSign, -CInt(expValue), CInt(expValue))  ' 3
            End If

            ' 符号を考慮して数値を計算
            Dim dres = CDbl(res) * Math.Pow(10.0, -decfig)
            Return If(isPlus, dres, -dres)
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、浮動小数点数の整数部分を変換します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換された浮動小数点数の整数部分。</returns>
        Private Function ToFloatDecInt(iter As U8String.U8StringIterator) As Integer
            ' 文字を数値化
            Dim res As Integer = 0
            While iter.HasNext()
                With iter.Current
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        res = res * 10 + (iter.Current.Value.Raw0 - &H30)
                    Else
                        ' 数字以外の文字が来た場合は終了
                        Exit While
                    End If
                End With
                iter.MoveNext()
            End While

            Return res
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、NaN（Not a Number）を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたNaN（Not a Number）。</returns>
        ''' <remarks>
        ''' 符号がある場合は、符号に応じてNaNを計算します。
        ''' </remarks>
        Function ConvertToNan(iter As U8String.U8StringIterator) As Double
            ' 符号が来た場合は読み進める
            With iter.Current
                If .HasValue Then
                    Select Case .Value.Raw0
                        Case &H2D ' -
                            Return -Double.NaN
                        Case &H2B ' +
                            Return Double.NaN
                    End Select
                End If
            End With

            ' 符号がない場合はNanを返す
            Return Double.NaN
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、Inf（Infinity）を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたInf（Infinity）。</returns>
        ''' <remarks>
        ''' 符号がある場合は、符号に応じてInfinityを計算します。
        ''' </remarks>
        Function ConvertToInf(iter As U8String.U8StringIterator) As Double
            ' 符号が来た場合は読み進める
            With iter.Current
                If .HasValue Then
                    Select Case .Value.Raw0
                        Case &H2D ' -
                            Return Double.NegativeInfinity
                        Case &H2B ' +
                            Return Double.PositiveInfinity
                    End Select
                End If
            End With

            ' 符号がない場合はInfを返す
            Return Double.PositiveInfinity
        End Function

#End Region

#Region "日時コンバート"

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、オフセット付き日時を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたオフセット付き日時。</returns>
        ''' <remarks>
        ''' 日時は、YYYY-MM-DDTHH:MM:SS[.SSS][Z|±HH:MM]の形式で表されます。
        ''' </remarks>
        Function ConvertToOffsetDateTime(iter As U8String.U8StringIterator) As DateTimeOffset
            ' 日付の部分を読み取る
            Dim dt = ReadDateBlock(iter)
            iter.MoveNext()

            ' 時間の部分を読み取る
            Dim tm = ReadTimeBlock(iter)

            Dim offTime As TimeSpan = TimeSpan.Zero
            If iter.HasNext() Then
                Select Case iter.Current.Value.Raw0
                    Case &H5A
                        ' Zulu time（UTC）を指定、オフセットはゼロ

                    Case &H2D
                        ' マイナスオフセット
                        iter.MoveNext()
                        With ReadSimpleTimeBlock(iter)
                            iter.MoveNext()
                            offTime = TimeSpan.FromMinutes(-(.hours * 60 + .minutes))
                        End With

                    Case &H2B
                        ' プラスオフセット
                        iter.MoveNext()
                        With ReadSimpleTimeBlock(iter)
                            iter.MoveNext()
                            offTime = TimeSpan.FromMinutes(.hours * 60 + .minutes)
                        End With
                End Select
            End If
            Return New DateTimeOffset(dt.year, dt.month, dt.day, tm.hours, tm.minutes, tm.seconds, tm.milliseconds, offTime)
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、ローカル日時を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたローカル日時。</returns>
        ''' <remarks>
        ''' 日時は、YYYY-MM-DDTHH:MM:SS[.SSS]の形式で表されます。
        ''' </remarks>
        Function ConvertToLocalDateTime(iter As U8String.U8StringIterator) As DateTime
            ' 日付の部分を読み取る
            Dim dt = ReadDateBlock(iter)
            iter.MoveNext()

            ' 時間の部分を読み取る
            Dim tm = ReadTimeBlock(iter)

            Return New DateTime(dt.year, dt.month, dt.day, tm.hours, tm.minutes, tm.seconds, tm.milliseconds, DateTimeKind.Local)
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、ローカル日付を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたローカル日付。</returns>
        ''' <remarks>
        ''' 日付は、YYYY-MM-DDの形式で表されます。
        ''' </remarks>
        Function ConvertToLocalDate(iter As U8String.U8StringIterator) As DateTime
            With ReadDateBlock(iter)
                Return New DateTime(.year, .month, .day, 0, 0, 0, DateTimeKind.Local)
            End With
        End Function

        ''' <summary>
        ''' 引数で指定されたU8Stringのイテレータから、ローカル時間を変換します。
        ''' </summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>変換されたローカル時間。</returns>
        ''' <remarks>
        ''' 時間は、時:分:秒[.SSS]の形式で表されます。
        ''' </remarks>
        Function ConvertToLocalTime(iter As U8String.U8StringIterator) As TimeSpan
            With ReadTimeBlock(iter)
                Return New TimeSpan(0, .hours, .minutes, .seconds, .milliseconds)
            End With
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、年月日ブロック（YYYY-MM-DD）を読み取ります。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>年月日のタプル（年、月、日）。</returns>
        ''' <remarks>
        ''' 年月日は4桁の年、2桁の月、2桁の日で構成されます。
        ''' </remarks>
        Private Function ReadDateBlock(iter As U8String.U8StringIterator) As (year As Integer, month As Integer, day As Integer)
            ' 年月日の部分を読み取る
            Dim year = ToTimeDecInt(iter, 4)
            iter.MoveNext()
            Dim month = ToTimeDecInt(iter, 2)
            iter.MoveNext()
            Dim day = ToTimeDecInt(iter, 2)
            Return (year, month, day)
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、時分ブロック（HH:MM）を読み取ります。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>時分のタプル（時間、分）。</returns>
        ''' <remarks>
        ''' 時分は2桁の時間と2桁の分で構成されます。
        ''' </remarks>
        Private Function ReadSimpleTimeBlock(iter As U8String.U8StringIterator) As (hours As Integer, minutes As Integer)
            ' 時分秒の部分を読み取る
            Dim hours = ToTimeDecInt(iter, 2)
            iter.MoveNext()
            Dim minutes = ToTimeDecInt(iter, 2)

            Return (hours, minutes)
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、時間ブロック（HH:MM:SS[.SSS]）を読み取ります。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <returns>時間ブロックのタプル（時間、分、秒、ミリ秒）。</returns>
        ''' <remarks>
        ''' ミリ秒は1000で割られた値として返されます。
        ''' </remarks>
        Private Function ReadTimeBlock(iter As U8String.U8StringIterator) As (hours As Integer, minutes As Integer, seconds As Integer, milliseconds As Integer)
            ' 時分秒の部分を読み取る
            Dim hours = ToTimeDecInt(iter, 2)
            iter.MoveNext()
            Dim minutes = ToTimeDecInt(iter, 2)
            iter.MoveNext()
            Dim seconds = ToTimeDecInt(iter, 2)

            ' あればミリ秒の部分を読み取る
            Dim milliseconds = 0
            If iter.HasNext() AndAlso EqualCharByte(iter.Current, &H2E) Then
                iter.MoveNext()
                milliseconds = ToTimeDecInt(iter, Integer.MaxValue)
                If milliseconds >= 1000 Then
                    milliseconds = milliseconds \ 1000
                End If
            End If

            Return (hours, minutes, seconds, milliseconds)
        End Function

        ''' <summary>引数で指定されたU8Stringのイテレータから、文字を整数に変換します。</summary>
        ''' <param name="iter">U8Stringのイテレータ。</param>
        ''' <param name="strlen">変換する文字列の長さ。</param>
        ''' <returns>文字列の整数値。</returns>
        Private Function ToTimeDecInt(iter As U8String.U8StringIterator, strlen As Integer) As Integer
            ' 文字を数値化
            Dim res As Integer = 0
            Dim len As Integer = 0
            While len < strlen AndAlso iter.HasNext()
                With iter.Current
                    If RangeCharByte(iter.Current, &H30, &H39) Then ' 0-9の範囲
                        res = res * 10 + (iter.Current.Value.Raw0 - &H30)
                    Else
                        ' 数字以外の文字が来た場合は終了
                        Exit While
                    End If
                End With
                iter.MoveNext()
                len += 1
            End While

            Return res
        End Function

#End Region

    End Module

End Namespace
