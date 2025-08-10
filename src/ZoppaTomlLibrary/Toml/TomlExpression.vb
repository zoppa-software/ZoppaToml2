Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>
    ''' TOMLの式を表す構造体。
    ''' この構造体は、TOMLの式の種類とその文字列を保持します。
    ''' </summary>
    ''' <param name="Type">式の種類を示す列挙型。</param>
    ''' <param name="Str">式を表す文字列。</param>
    Public Structure TomlExpression

        ' 空のインスタンスを生成するためのLazy初期化
        Private Shared ReadOnly _empty As New Lazy(Of TomlExpression)(Function() New TomlExpression(TomlExpressionType.None, U8String.Empty))

        ''' <summary>
        ''' 空のTomlExpressionを取得します。
        ''' </summary>
        ''' <remarks>
        ''' このプロパティは、空のTomlExpressionインスタンスを返します。
        ''' 空のインスタンスは、式が存在しないことを示すために使用されます。
        ''' </remarks>
        Public Shared ReadOnly Property Empty As TomlExpression
            Get
                Return _empty.Value
            End Get
        End Property

        ''' <summary>
        ''' 式の内容を表すTomlExpressionの配列。
        ''' </summary>
        ''' <remarks>
        ''' このフィールドは、式が複数のトークンから構成される場合に使用されます。
        ''' </remarks>

        ' 子要素
        Private _contents As TomlExpression()

        ''' <summary>式の種類を取得します。</summary>
        ''' <remarks>この列挙型は、TOMLの式の種類を定義します。</remarks>
        Public ReadOnly Property Type As TomlExpressionType

        ''' <summary>式を表す文字列を取得します。</summary>
        ''' <remarks>このプロパティは、TOMLの式を表す文字列を保持します。</remarks>
        ''' <returns>式を表すUTF-8文字列。</returns>
        Public ReadOnly Property Str As U8String

        ''' <summary>
        ''' 式の内容を取得します。
        ''' </summary>
        ''' <returns>式の内容を表すTomlExpressionの配列。</returns>
        ''' <remarks>
        ''' このプロパティは、式の内容を表すTomlExpressionの配列を返します。
        ''' 例えば、式が複数のトークンから構成される場合に使用されます。
        ''' </remarks>
        Public ReadOnly Property Contents As TomlExpression()
            Get
                Return _contents
            End Get
        End Property

        ''' <summary>
        ''' TomlExpressionの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="type">式の種類を示す列挙型。</param>
        ''' <param name="str">式を表す文字列。</param>
        ''' <remarks>
        ''' このコンストラクタは、指定された式の種類と文字列でTomlExpressionの新しいインスタンスを作成します。
        ''' </remarks>
        Public Sub New(type As TomlExpressionType, str As U8String)
            Me.Type = type
            Me.Str = str
            Me._contents = New TomlExpression() {}
        End Sub

        ''' <summary>
        ''' TomlExpressionの新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="type">式の種類を示す列挙型。</param>
        ''' <param name="str">式を表す文字列。</param>
        ''' <param name="contents">式の内容を表すTomlExpressionの配列。</param>
        ''' <remarks>
        ''' このコンストラクタは、指定された式の種類、文字列、および内容でTomlExpressionの新しいインスタンスを作成します。
        ''' </remarks>
        Public Sub New(type As TomlExpressionType, str As U8String, contents As TomlExpression())
            Me.Type = type
            Me.Str = str
            Me._contents = contents
        End Sub

        ''' <summary>
        ''' 式の内容を設定します。
        ''' 
        ''' このメソッドは、式の内容を表すTomlExpressionの配列を設定します。
        ''' 例えば、式が複数のトークンから構成される場合に使用されます。
        ''' </summary>
        ''' <param name="contents">式の内容。</param>
        Public Sub SetContents(contents As TomlExpression())
            Me._contents = contents
        End Sub

        ''' <summary>
        ''' 式の内容を文字列として返します。
        ''' </summary>
        ''' <returns>式を表す文字列。</returns>
        ''' <remarks>
        ''' このメソッドは、式の内容を文字列として返します。
        ''' 例えば、式が複数のトークンから構成される場合に使用されます。
        ''' </remarks>
        Public Overrides Function ToString() As String
            Return $"{Type} : {Str}"
        End Function

        ''' <summary>
        ''' 式の内容を指定した型に変換します。
        ''' このメソッドは、式の内容を指定された型に変換します。
        ''' 例えば、基本文字列やマルチライン基本文字列に変換する場合に使用されます。
        ''' </summary>
        ''' <typeparam name="T">型。</typeparam>
        ''' <returns>値。</returns>
        Public Function ValueTo(Of T)() As T
            Select Case Me.Type
                Case TomlExpressionType.BasicString
                    Return CType(CObj(ConvertToBasicString(Me.Str.GetIterator()).ToString()), T)
                Case TomlExpressionType.MlBasicString
                    Return CType(CObj(ConvertToMlBasicString(Me.Str.GetIterator()).ToString()), T)
                Case TomlExpressionType.LiteralString
                    Return CType(CObj(ConvertToLiteralString(Me.Str.GetIterator()).ToString()), T)
                Case TomlExpressionType.MlLiteralString
                    Return CType(CObj(ConvertToMlLiteralString(Me.Str.GetIterator()).ToString()), T)
                Case TomlExpressionType.DecInt
                    Return CType(CObj(ConvertToDecInt(Me.Str.GetIterator())), T)
                Case TomlExpressionType.HexInt
                    Return CType(CObj(ConvertToDecInt(Me.Str.GetIterator(), 4)), T)
                Case TomlExpressionType.OctInt
                    Return CType(CObj(ConvertToDecInt(Me.Str.GetIterator(), 3)), T)
                Case TomlExpressionType.BinInt
                    Return CType(CObj(ConvertToDecInt(Me.Str.GetIterator(), 1)), T)
                Case TomlExpressionType.Float
                    Return CType(CObj(ConvertToFloat(Me.Str.GetIterator())), T)
                Case TomlExpressionType.Inf
                    Return CType(CObj(ConvertToInf(Me.Str.GetIterator())), T)
                Case TomlExpressionType.Nan
                    Return CType(CObj(ConvertToNan(Me.Str.GetIterator())), T)
                Case TomlExpressionType.BooleanLiteral
                    Return CType(CObj(Me.Str = U8String.NewString({&H74, &H72, &H75, &H65})), T)
                Case TomlExpressionType.OffsetDateTime
                    Return CType(CObj(ConvertToOffsetDateTime(Me.Str.GetIterator())), T)
                Case TomlExpressionType.LocalDateTime
                    Return CType(CObj(ConvertToLocalDateTime(Me.Str.GetIterator())), T)
                Case TomlExpressionType.LocalDate
                    Return CType(CObj(ConvertToLocalDate(Me.Str.GetIterator())), T)
                Case TomlExpressionType.LocalTime
                    Return CType(CObj(ConvertToLocalTime(Me.Str.GetIterator())), T)
                Case Else
                    Throw New InvalidCastException($"{Type}を{GetType(T)}に変換できません")
            End Select
        End Function

    End Structure

End Namespace
