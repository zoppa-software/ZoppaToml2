Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLの式の種類を定義する列挙型。
    ''' この列挙型は、TOMLの式の種類を示します。
    ''' </summary>
    Public Enum TomlExpressionType

        ''' <summary>式が存在しないことを示します。</summary>
        None

        ''' <summary>式。</summary>
        Expression

        ''' <summary>空白。</summary>
        Ws

        ''' <summary>コメント。</summary>
        Comment

        ''' <summary>キー、値。</summary>
        Keyval

        ''' <summary>引用符なしのキー。</summary>
        UnquotedKey

        ''' <summary>簡単なキー。</summary>
        SimpleKey

        ''' <summary>ドット連結キー。</summary>
        DottedKey

        ''' <summary>引用符のキー。</summary>
        QuotedKey

        ''' <summary>ドット区切り。</summary>
        DotSep

        ''' <summary>等号区切り。</summary>
        KeyvalSep

        ''' <summary>基本文字列。</summary>
        BasicString

        ''' <summary>複数行基本文字列。</summary>
        MlBasicString

        ''' <summary>複数行リテラル文字列。</summary>
        MlLiteralString

        ''' <summary>リテラル文字列。</summary>
        LiteralString

        ''' <summary>符号つき整数。</summary>
        DecInt

        ''' <summary>16進数。</summary>
        HexInt

        ''' <summary>8進数。</summary>
        OctInt

        ''' <summary>2進数。</summary>
        BinInt

        ''' <summary>浮動小数点数。</summary>
        Float

        ''' <summary>TomlExpressionType.Inf</summary>
        Inf

        ''' <summary>NaN（Not a Number）</summary>
        Nan

        ''' <summary>真偽値。</summary>
        BooleanLiteral

        ''' <summary>オフセット日時。</summary>
        OffsetDateTime

        ''' <summary>ローカル日時。</summary>
        LocalDateTime

        ''' <summary>ローカル日付。</summary>
        LocalDate

        ''' <summary>ローカル時間。</summary>
        LocalTime

        ''' <summary>配列。</summary>
        Array

        ''' <summary>配列の値。</summary>
        ArrayValues

        ''' <summary>空白-コメント-改行。</summary>
        WsCommentNewline

        ''' <summary>標準テーブル。</summary>
        Table

        ''' <summary>インラインテーブル。</summary>
        InlineTable

        ''' <summary>インラインテーブル、キー値。</summary>
        InlineTableKeyvals

        ''' <summary>配列テーブル。</summary>
        ArrayTable

    End Enum

End Namespace
