Option Strict On
Option Explicit On

Namespace Strings

    ''' <summary>
    ''' 文字列（UTF-8）
    ''' </summary>
    Public Structure U8String
        Implements IComparable(Of U8String), IEquatable(Of U8String)

        ' 参照開始位置
        Private ReadOnly _start As Integer

        ' 参照終了位置
        Private ReadOnly _end As Integer

        ' 元のバイト配列を取得します。
        Private ReadOnly raw() As Byte

        ''' <summary>文字列の文字数を取得します。</summary>
        ''' <returns>文字列の長さ。</returns>
        Public ReadOnly Property Length As Integer

        ''' <summary>参照範囲のバイト数を取得します。</summary>
        ''' <returns>参照範囲のバイト数。</returns>
        ''' <remarks>ByteLengthプロパティは、文字数ではなくバイト数を返します。</remarks>
        Public ReadOnly Property ByteLength As Integer
            Get
                Return Me._end - Me._start
            End Get
        End Property

        ''' <summary>
        ''' 参照範囲のバイト配列を取得します。
        ''' 参照範囲内のバイトデータを返します。
        ''' </summary>
        ''' <returns>参照範囲のバイト配列。</returns>
        Public ReadOnly Property Data As Byte()
            Get
                If Me.raw Is Nothing Then
                    Return Nothing
                End If
                Dim res() As Byte = New Byte(Me._end - Me._start - 1) {}
                Array.Copy(Me.raw, Me._start, res, 0, Me._end - Me._start)
                Return res
            End Get
        End Property

        ''' <summary>
        ''' 空の文字列を取得します。
        ''' 参照範囲が空のU8Stringを返します。
        ''' </summary>
        ''' <returns>空のU8String。</returns>
        ''' <remarks>Emptyプロパティは、初期化されていない状態のU8Stringを表します。</remark
        Public Shared ReadOnly Property Empty As U8String
            Get
                Return New U8String(0, 0, Nothing)
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="start">参照開始位置。</param>
        ''' <param name="[end]">参照終了位置。</param>
        ''' <param name="source">元配列。</param>
        Private Sub New(start As Integer, [end] As Integer, source() As Byte)
            If source Is Nothing Then
                source = New Byte() {}
            End If
            If start < 0 OrElse [end] < start OrElse [end] > source.Length Then
                Throw New ArgumentOutOfRangeException("有効な範囲を指定していません")
            End If
            Me._start = start
            Me._end = [end]
            Me.raw = source
            Me.Length = Me.CountLendth()
        End Sub

        ''' <summary>
        ''' 文字列（UTF-8）を文字列から取得します。
        ''' </summary>
        ''' <returns>文字列。</returns>
        Public Shared Function NewString(source As String) As U8String
            If source Is Nothing Then
                Throw New ArgumentNullException("source", "文字列はnullにできません")
            End If
            Dim bytes() As Byte = System.Text.Encoding.UTF8.GetBytes(source)
            Return New U8String(0, bytes.Length, bytes)
        End Function

        ''' <summary>
        ''' 文字列（UTF-8）をバイト配列から取得します。
        ''' バイト配列の所有者はこのU8Stringに移動されません。
        ''' </summary>
        ''' <returns>文字列。</returns>
        Public Shared Function NewString(bytes As Byte()) As U8String
            If bytes Is Nothing Then
                Throw New ArgumentNullException("source", "バイト配列はnullにできません")
            End If
            Dim newbyte() As Byte = New Byte(bytes.Length - 1) {}
            Array.Copy(bytes, 0, newbyte, 0, bytes.Length)
            Return New U8String(0, newbyte.Length, newbyte)
        End Function

        ''' <summary>
        ''' 文字列（UTF-8）をバイト配列から取得します。
        ''' バイト配列の所有者はこのU8Stringに移動されます。
        ''' このメソッドは、バイト配列を直接参照するため、元のバイト配列は変更されないことに注意してください。
        ''' </summary>
        ''' <returns>文字列。</returns>
        Public Shared Function NewStringChangeOwner(bytes As Byte()) As U8String
            If bytes Is Nothing Then
                Throw New ArgumentNullException("source", "バイト配列はnullにできません")
            End If
            Return New U8String(0, bytes.Length, bytes)
        End Function

        ''' <summary>
        ''' 文字列（UTF-8）を範囲を指定して取得します。
        ''' </summary>
        ''' <returns>文字列。</returns>
        Public Shared Function NewSlice(source As U8String, start As Integer, length As Integer) As U8String
            If source.raw Is Nothing Then
                Throw New ArgumentNullException("source", "元の文字列はnullにできません")
            End If
            Return New U8String(source._start + start, source._start + start + length, source.raw)
        End Function

        ''' <summary>
        ''' 全体の文字数を取得します。
        ''' 参照範囲内のUTF-8文字の数を返します。
        ''' </summary>
        ''' <returns>文字文字数。</returns>
        Private Function CountLendth() As Integer
            Dim res As Integer = 0
            Dim i As Integer = Me._start
            While i < Me._end
                i += U8Char.Utf8ByteSequenceLength(Me.raw(i))
                res += 1
            End While
            Return res
        End Function

        ''' <summary>参照範囲のバイト配列をクローンし、新しくメモリを割り当ててU8Stringとして返します。</summary>
        ''' <returns>新しいU8String。</returns>
        Public Function NewAllocate() As U8String
            Return U8String.NewStringChangeOwner(Me.Data)
        End Function

        ''' <summary>
        ''' 文字列を連結します。
        ''' 参照範囲の文字列と指定された文字列を連結して新しいU8Stringを返します。
        ''' </summary>
        ''' <param name="other">連結する文字列。</param>
        ''' <returns>連結された新しいU8String。</returns>
        ''' <remarks>このメソッドは、元の文字列を変更しません。</remarks>
        Public Function Concat(other As U8String) As U8String
            If other.raw Is Nothing Then
                Return Me
            ElseIf Me.raw Is Nothing Then
                Return other
            End If

            Dim buffer As Byte() = New Byte(Me._end - Me._start + other._end - other._start - 1) {}
            Array.Copy(Me.raw, Me._start, buffer, 0, Me._end - Me._start)
            Array.Copy(other.raw, other._start, buffer, Me._end - Me._start, other._end - other._start)
            Return New U8String(0, buffer.Length, buffer)
        End Function

        ''' <summary>
        ''' 文字列を連結します。
        ''' 参照範囲の文字列と指定された文字列を連結して新しいU8Stringを返します。
        ''' </summary>
        ''' <param name="other">連結する文字列。</param>
        ''' <returns>連結された新しいU8String。</returns>
        ''' <remarks>このメソッドは、元の文字列を変更しません。</remarks>
        Public Function Concat(other As String) As U8String
            Return Me.Concat(U8String.NewString(other))
        End Function

        ''' <summary>
        ''' 文字列の一部を取得します。
        ''' 指定された開始位置から指定された長さの文字列を返します。
        ''' </summary>
        ''' <param name="first">開始位置（0から始まるインデックス）。</param>
        ''' <param name="length">取得する文字数。</param>
        ''' <returns>指定された範囲の文字列。</returns>
        Public Function Mid(first As Integer, length As Integer) As U8String
            If first < 0 OrElse length < 0 OrElse first + length > Me.Length Then
                Throw New ArgumentOutOfRangeException("範囲外の引数が指定されました")
            End If

            Dim charRange = Me.GetRange(first, first + length)
            Return New U8String(charRange.byteStart, charRange.byteEnd, Me.raw)
        End Function

        ''' <summary>
        ''' 文字列の範囲を取得します。
        ''' 指定された開始位置と終了位置の範囲内の文字列を返します。
        ''' </summary>
        ''' <param name="first">開始位置（0から始まるインデックス）。</param>
        ''' <param name="last">終了位置（0から始まるインデックス）。</param>
        ''' <returns>指定された範囲の文字列。</returns>
        Private Function GetRange(first As Integer, last As Integer) As (byteStart As Integer, byteEnd As Integer, charCount As Integer)
            Dim idxs() As Integer = {first, last}
            Dim res() As Integer = {Me.raw.Length, Me.raw.Length, 0}

            Dim pos As Integer = 0
            Dim i As Integer = Me._start, j As Integer = 0
            While i < Math.Min(Me._end, Me.raw.Length)
                ' 現在のバイトのUTF-8シーケンスの長さを取得します
                Dim ln As Integer = U8Char.Utf8ByteSequenceLength(Me.raw(i))

                ' 現在の位置が開始位置または終了位置に一致するか確認します
                If pos = idxs(j) Then
                    ' 開始位置または終了位置に一致した場合、結果に追加します
                    res(j) = i
                    j += 1
                    If j >= 2 Then
                        ' 両方の位置が見つかった場合、ループを終了します
                        Exit While
                    End If
                Else
                    pos += 1
                    i += ln
                End If
            End While

            res(2) = Math.Min(pos, last) - first
            Return (res(0), res(1), res(2))
        End Function

        ''' <summary>
        ''' 指定されたインデックスの文字を取得します。
        ''' 参照範囲内のUTF-8文字を指定されたインデックスで取得します。
        ''' </summary>
        ''' <param name="index">取得する文字のインデックス（0から始まる）。</param>
        ''' <returns>指定されたインデックスの文字。範囲外の場合はNothing。</returns>
        ''' <remarks>インデックスが範囲外の場合はNothingを返します。</remarks>
        Public Function At(index As Integer) As U8Char?
            If index < 0 OrElse index >= Me.Length Then
                Return Nothing
            End If
            Dim pos As Integer = 0
            Dim i As Integer = Me._start
            While i < Me._end
                Dim charLength As Integer = U8Char.Utf8ByteSequenceLength(Me.raw(i))
                If pos = index Then
                    Return U8Char.NewChar(Me.raw, i)
                End If
                pos += 1
                i += charLength
            End While
            Return Nothing ' 指定されたインデックスが範囲外の場合はNothingを返す
        End Function

        ''' <summary>
        ''' 文字列が等しいかどうかを確認します。
        ''' 参照範囲内の文字列が、指定された文字列と等しい場合はTrueを返します。
        ''' </summary>
        ''' <param name="obj">比較するオブジェクト。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定された文字列と等しいです。</returns>
        Public Overrides Function Equals(obj As Object) As Boolean
            If TypeOf obj Is U8String Then
                Return Equals(CType(obj, U8String))

            ElseIf TypeOf obj Is String Then
                ' 文字列をUTF-8バイト配列に変換して比較
                Dim other As String = CType(obj, String)
                Dim otherBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(other)
                If otherBytes.Length <> (Me._end - Me._start) Then
                    Return False
                End If
                For i As Integer = 0 To otherBytes.Length - 1
                    If Me.raw(Me._start + i) <> otherBytes(i) Then
                        Return False
                    End If
                Next
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 文字列が等しいかどうかを確認します。
        ''' 参照範囲内の文字列が、指定されたU8Stringと等しい場合はTrueを返します。
        ''' </summary>
        ''' <param name="other">比較するU8String。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定されたU8Stringと等しいです。</returns>
        Public Overloads Function Equals(other As U8String) As Boolean Implements IEquatable(Of U8String).Equals
            ' 参照範囲の長さが異なる場合は等しくない
            If (Me._end - Me._start) <> (other._end - other._start) Then
                Return False
            End If

            ' バイトごとに比較
            For i As Integer = 0 To Me._end - Me._start - 1
                If Me.raw(Me._start + i) <> other.raw(other._start + i) Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' = 演算子をオーバーロードします。
        ''' 2つのU8Charが等しいかどうかを比較します。
        ''' </summary>
        ''' <returns>等しければ真。</returns>
        Public Overloads Shared Operator =(ByVal c1 As U8String, ByVal c2 As U8String) As Boolean
            If (c1._end - c1._start) <> (c2._end - c2._start) Then
                Return False
            End If
            ' バイトごとに比較
            For i As Integer = 0 To c1._end - c1._start - 1
                If c1.raw(c1._start + i) <> c2.raw(c2._start + i) Then
                    Return False
                End If
            Next
            Return True
        End Operator

        ''' <summary>
        ''' <> 演算子をオーバーロードします。
        ''' 2つのU8Charが等しくないかどうかを比較します。
        ''' </summary>
        ''' <returns>等しくなければ真。</returns>
        Public Overloads Shared Operator <>(ByVal c1 As U8String, ByVal c2 As U8String) As Boolean
            If (c1._end - c1._start) <> (c2._end - c2._start) Then
                Return True
            End If
            ' バイトごとに比較
            For i As Integer = 0 To c1._end - c1._start - 1
                If c1.raw(c1._start + i) <> c2.raw(c2._start + i) Then
                    Return True
                End If
            Next
            Return False
        End Operator

        ''' <summary>
        ''' 文字列を比較します。
        ''' 参照範囲内の文字列と指定された文字列を比較し、辞書順での大小を返します。
        ''' </summary>
        ''' <param name="other">比較する文字列。</param>
        ''' <returns>0ならば等しい、負の値ならば小さい、正の値ならば大きい。</returns>
        Public Function CompareTo(other As U8String) As Integer Implements IComparable(Of U8String).CompareTo
            ' 文字列の長さを比較します。
            Dim selfLen = Me.Length
            Dim otherLen = other.Length
            Dim minLen = Math.Min(selfLen, otherLen)

            ' 文字列のバイト配列を比較します
            Dim sskip As Integer = 0
            Dim oskip As Integer = 0
            For i As Integer = 0 To minLen - 1
                Dim s1 = U8Char.NewChar(Me.raw, Me._start + sskip)
                Dim o1 = U8Char.NewChar(other.raw, other._start + oskip)

                If s1 <> o1 Then
                    Return If(s1.IntegerValue < o1.IntegerValue, -1, 1)
                End If

                sskip += s1.Size
                oskip += o1.Size
            Next

            ' 文字列の長さを比較します
            ' 長さが異なる場合は、長い方が大きいと判断します
            If selfLen <> otherLen Then
                Return If(selfLen < otherLen, -1, 1)
            End If
            Return 0
        End Function

        ''' <summary>
        ''' 文字列が他の文字列で始まるかどうかを確認します。
        ''' 参照範囲内の文字列が、指定された文字列で始まる場合はTrueを返します。
        ''' </summary>
        ''' <param name="other">比較する文字列。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定された文字列で始まります。</returns>
        Public Function StartWith(other As U8String) As Boolean
            ' 比較する文字列がnullまたは空の場合はFalseを返す
            If other.raw Is Nothing OrElse Me.raw Is Nothing Then
                Return False
            End If
            ' 比較する文字列の長さが参照範囲より大きい場合はFalseを返す
            If other._end - other._start > Me._end - Me._start Then
                Return False
            End If
            ' 比較する文字が空白なら一致としない
            If other._end - other._start <= 0 OrElse Me._end - Me._start <= 0 Then
                Return False
            End If
            ' 比較範囲の文字列を1文字ずつ比較する
            For i As Integer = 0 To other._end - other._start - 1
                If Me.raw(Me._start + i) <> other.raw(other._start + i) Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' 文字列が他の文字列で始まるかどうかを確認します。
        ''' 参照範囲内の文字列が、指定された文字列で始まる場合はTrueを返します。
        ''' </summary>
        ''' <param name="other">比較する文字列。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定された文字列で始まります。</returns>
        Public Function StartWith(other As String) As Boolean
            Return Me.StartWith(U8String.NewString(other))
        End Function

        ''' <summary>
        ''' 文字列が他の文字列で終わるかどうかを確認します。
        ''' 参照範囲内の文字列が、指定された文字列で終わる場合はTrueを返します。
        ''' </summary>
        ''' <param name="other">比較する文字列。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定された文字列で終わります。</returns>
        Public Function EndWith(other As U8String) As Boolean
            ' 比較する文字列がnullまたは空の場合はFalseを返す
            If other.raw Is Nothing OrElse Me.raw Is Nothing Then
                Return False
            End If
            ' 比較する文字列の長さが参照範囲より大きい場合はFalseを返す
            If other._end - other._start > Me._end - Me._start Then
                Return False
            End If
            ' 比較する文字が空白なら一致としない
            If other._end - other._start <= 0 OrElse Me._end - Me._start <= 0 Then
                Return False
            End If
            ' 参照範囲の末尾から比較する
            Dim startIdx As Integer = Me._end - other._end + other._start
            For i As Integer = 0 To other._end - other._start - 1
                If Me.raw(startIdx + i) <> other.raw(other._start + i) Then
                    Return False
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' 文字列が他の文字列で終わるかどうかを確認します。
        ''' 参照範囲内の文字列が、指定された文字列で終わる場合はTrueを返します。
        ''' </summary>
        ''' <param name="other">比較する文字列。</param>
        ''' <returns>Trueならば、参照範囲の文字列は指定された文字列で終わります。</returns>
        Public Function EndWith(other As String) As Boolean
            Return Me.EndWith(U8String.NewString(other))
        End Function

        ''' <summary>
        ''' 文字列の先頭と末尾の空白を削除します。
        ''' 参照範囲内の文字列から先頭と末尾の空白文字を削除した新しいU8Stringを返します。
        ''' 参照範囲が空の場合は空のU8Stringを返します。
        ''' </summary>
        ''' <returns>トリムした文字列。</returns>
        Public Function Trim() As U8String
            Dim st As Integer = Me._start
            Dim ed As Integer = Me._end
            Dim sk As Integer = 0

            ' 先頭の空白をスキップ
            While st < ed
                Dim c As U8Char
                Select Case U8Char.Utf8ByteSequenceLength(Me.raw(st))
                    Case 1
                        c = New U8Char(Me.raw(st))
                        sk = 1
                    Case 2
                        c = New U8Char(Me.raw(st), Me.raw(st + 1))
                        sk = 2
                    Case 3
                        c = New U8Char(Me.raw(st), Me.raw(st + 1), Me.raw(st + 2))
                        sk = 3
                    Case 4
                        c = New U8Char(Me.raw(st), Me.raw(st + 1), Me.raw(st + 2), Me.raw(st + 3))
                        sk = 4
                End Select
                If c.IsWhiteSpace Then
                    st += sk
                Else
                    Exit While
                End If
            End While

            ' 末尾の空白をスキップ
            While ed > st
                Dim c As U8Char
                If (Me.raw(ed - 1) And &HC0) <> &H80 Then
                    c = New U8Char(Me.raw(ed - 1))
                    sk = 1
                ElseIf (Me.raw(ed - 2) And &HC0) <> &H80 Then
                    c = New U8Char(Me.raw(ed - 2), Me.raw(ed - 1))
                    sk = 2
                ElseIf (Me.raw(ed - 3) And &HC0) <> &H80 Then
                    c = New U8Char(Me.raw(ed - 3), Me.raw(ed - 2), Me.raw(ed - 1))
                    sk = 3
                ElseIf (Me.raw(ed - 4) And &HC0) <> &H80 Then
                    c = New U8Char(Me.raw(ed - 4), Me.raw(ed - 3), Me.raw(ed - 2), Me.raw(ed - 1))
                    sk = 4
                End If
                If c.IsWhiteSpace Then
                    ed -= sk
                Else
                    Exit While
                End If
            End While

            ' 新しいU8Stringを返す
            If st >= ed Then
                Return U8String.Empty ' 空の文字列を返す
            End If
            Return New U8String(st, ed, Me.raw)
        End Function

        ''' <summary>
        ''' 文字イテレーターを取得します。
        ''' 参照範囲内のUTF-8文字を1つずつ返すイテレーターを返します。
        ''' </summary>
        ''' <returns>イテレーター。</returns>
        Public Function GetIterator() As U8StringIterator
            Return New U8StringIterator(Me.raw, Me._start, Me._end)
        End Function

        ''' <summary>
        ''' 文字列のバイトイテレーターを取得します。
        ''' 参照範囲内のバイトを1つずつ返すイテレーターを返します。
        ''' </summary>
        ''' <returns>バイトイテレーター。</returns>
        Public Function GetByteEnumerator() As U8StringByteEnumerator
            Return New U8StringByteEnumerator(Me)
        End Function

        ''' <summary>
        ''' 文字列のハッシュコードを取得します。
        ''' 参照範囲内のバイト配列を基にハッシュコードを計算します。
        ''' </summary>
        ''' <returns>ハッシュコード。</returns>

        ''' <summary>
        ''' 文字列表現を取得します。
        ''' 参照範囲内のバイト配列をUTF-8文字列に変換して返します。
        ''' 参照範囲が空の場合は空文字列を返します。
        ''' </summary>
        ''' <returns>文字列表現。</returns>
        Public Overrides Function ToString() As String
            If Me.raw Is Nothing OrElse Me._start >= Me._end Then
                Return String.Empty
            End If
            Return System.Text.Encoding.UTF8.GetString(Me.raw, Me._start, Me._end - Me._start)
        End Function

        ''' <summary>
        ''' 文字列のイテレーター。
        ''' 参照範囲内のUTF-8文字を1つずつ返すイテレーターです。
        ''' </summary>
        Public NotInheritable Class U8StringIterator

            ' 元のバイト配列
            Private ReadOnly raw As Byte()

            ' 参照開始位置
            Private ReadOnly _start As Integer

            ' 参照終了位置
            Private ReadOnly _end As Integer

            ' 参照位置
            Private _current As Integer

            ''' <summary>
            ''' 次の文字があるかどうかを取得します。
            ''' 参照範囲内に次の文字が存在する場合はTrueを返します。
            ''' </summary>
            ''' <returns>文字が存在したら真を返す。</returns>
            Public ReadOnly Property HasNext As Boolean
                Get
                    Return Me._start + Me._current < Me._end
                End Get
            End Property

            ''' <summary>現在のインデックスを取得します。</summary>
            ''' <returns>現在のインデックス。</returns>
            ''' <remarks>イテレーターの状態を更新しません。</remarks>
            Public ReadOnly Property CurrentIndex() As Integer
                Get
                    Return Me._current
                End Get
            End Property

            ''' <summary>
            ''' 現在の文字を取得します。
            ''' 参照位置が有効な範囲内であれば、現在のUTF-8文字を返します。
            ''' それ以外の場合はNothingを返します。
            ''' </summary>
            ''' <returns></returns>
            Public ReadOnly Property Current As U8Char?
                Get
                    If Me._current < Me._end - Me._start Then
                        Return U8Char.NewChar(Me.raw, Me._start + Me._current)
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="raw">元のバイト配列。</param>
            ''' <param name="start">参照開始位置。</param>
            ''' <param name="[end]">参照終了位置。</param>
            Public Sub New(raw() As Byte, start As Integer, [end] As Integer)
                Me.raw = raw
                Me._start = start
                Me._end = [end]
                Me._current = 0
            End Sub

            ''' <summary>
            ''' 次の文字を取得します。
            ''' 参照位置を次の文字に進め、現在の文字を返します。
            ''' 参照位置が範囲外の場合はNothingを返します。
            ''' </summary>
            ''' <returns>次の文字。</returns>
            ''' <remarks>イテレーターの状態を更新します。</remarks>
            Public Function MoveNext() As U8Char?
                Dim res = Me.Current
                If res IsNot Nothing Then
                    Me._current += res.Value.Size
                End If
                Return res
            End Function

            ''' <summary>
            ''' 指定した数の文字をスキップします。
            ''' 参照位置を指定された数だけ進め、スキップ後の文字を返します。
            ''' </summary>
            ''' <param name="skipCount">スキップする文字数。</param>
            ''' <returns>スキップ後の文字。</returns>
            ''' <remarks>イテレーターの状態を更新しません。</remarks>
            Public Function Peek(skipCount As Integer) As U8Char?
                Dim ci = Me._current
                Dim i As Integer = 0
                Dim ln As Integer
                While i < skipCount
                    If Me._start + ci < Me.raw.Length Then
                        ln = U8Char.Utf8ByteSequenceLength(Me.raw(Me._start + ci))
                        i += 1
                        ci += ln
                    Else
                        Return Nothing
                    End If
                End While
                If Me._start + ci < Me.raw.Length Then
                    Return U8Char.NewChar(Me.raw, Me._start + ci)
                Else
                    Return Nothing
                End If
            End Function

            ''' <summary>
            ''' イテレーターをクローンします。
            ''' 現在のイテレーターの状態を保持した新しいイテレーターを返します。
            ''' </summary>
            ''' <returns>新しいイテレーター。</returns>
            Public Function Clone() As U8StringIterator
                Return New U8StringIterator(Me.raw, Me._start, Me._end) With {
                    ._current = Me._current
                }
            End Function

            ''' <summary>
            ''' カレント位置を設定します。
            ''' イテレーターの現在の位置を指定されたインデックスに設定します。
            ''' </summary>
            ''' <param name="current">新しい位置。</param>
            Sub SetCurrentIndex(current As Integer)
                _current = current
            End Sub

        End Class

        ''' <summary>
        ''' 文字列のバイトイテレーター。
        ''' 参照範囲内のバイトを1つずつ返すイテレーターです。
        ''' </summary>
        Public NotInheritable Class U8StringByteEnumerator
            Implements IEnumerable(Of Byte)

            Private ReadOnly _target As U8String

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="target">U8Stringオブジェクト。</param>
            Public Sub New(target As U8String)
                Me._target = target
            End Sub

            ''' <summary>イテレーターを取得します。</summary>
            ''' <returns>Byteの列挙子。</returns>
            Public Function GetEnumerator() As IEnumerator(Of Byte) Implements IEnumerable(Of Byte).GetEnumerator
                Return New U8StringByteIterator(_target.raw, _target._start, _target._end)
            End Function

            ''' <summary>非ジェネリックなイテレーターを取得します。</summary>
            ''' <returns>非ジェネリックな列挙子。</returns>
            Private Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
                Return Me.GetEnumerator()
            End Function

        End Class

        ''' <summary>
        ''' 文字列のバイトイテレーター。
        ''' 参照範囲内のバイトを1つずつ返すイテレーターです。
        ''' </summary>
        Class U8StringByteIterator
            Implements IEnumerator(Of Byte)

            ' 元のバイト配列
            Private ReadOnly raw As Byte()

            ' 参照開始位置
            Private ReadOnly _start As Integer

            ' 参照終了位置
            Private ReadOnly _end As Integer

            ' 参照位置
            Private _current As Integer

            ''' <summary>コンストラクタ。</summary>
            ''' <param name="raw">元のバイト配列。</param>
            ''' <param name="start">参照開始位置。</param>
            ''' <param name="[end]">参照終了位置。</param>
            Public Sub New(raw() As Byte, start As Integer, [end] As Integer)
                Me.raw = raw
                Me._start = start
                Me._end = [end]
                Me._current = -1
            End Sub

            ''' <summary>
            ''' 現在のバイトを取得します。
            ''' 参照位置が有効な範囲内であれば、現在のバイトを返します。
            ''' それ以外の場合は例外をスローします。
            ''' </summary>
            ''' <returns>現在位置のバイト値。</returns>
            Public ReadOnly Property Current As Byte Implements IEnumerator(Of Byte).Current
                Get
                    Return Me.raw(Me._start + Me._current)
                End Get
            End Property

            ''' <summary>
            ''' 現在のバイトを取得します。
            ''' 参照位置が有効な範囲内であれば、現在のバイトを返します。
            ''' それ以外の場合は例外をスローします。
            ''' </summary>
            ''' <returns>現在位置のバイト値。</returns>
            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return Current
                End Get
            End Property

            ''' <summary>イテレーターをリセットします。</summary>
            Public Sub Reset() Implements IEnumerator.Reset
                Me._current = -1
            End Sub

            ''' <summary>
            ''' 次のバイトに進みます。
            ''' 参照位置を次のバイトに進め、次のバイトが範囲内であればTrueを返します。
            ''' </summary>
            ''' <returns>次のバイトが存在する場合はTrue。</returns>
            ''' <exception cref="InvalidOperationException">イテレーターの位置が範囲外です。</exception>
            Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
                Me._current += 1
                Return Me._current < Me._end - Me._start
            End Function

            ''' <summary>破棄処理を行います。</summary>
            ''' <remarks>このメソッドは、IDisposableインターフェイスを実装しています。</remarks>
            Public Sub Dispose() Implements IDisposable.Dispose
                ' 未使用
            End Sub

        End Class

    End Structure

End Namespace

