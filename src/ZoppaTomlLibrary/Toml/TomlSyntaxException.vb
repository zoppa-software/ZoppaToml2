Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLの構文エラーを表す例外クラス。
    ''' このクラスは、TOMLの解析中に構文エラーが発生した場合にスローされます。
    ''' </summary>
    ''' <remarks>
    ''' この例外は、TOMLの解析中に無効な構文が検出された場合にスローされます。
    ''' 例えば、無効なキーや値、誤ったテーブルの定義などが含まれます。
    ''' </remarks>
    Public NotInheritable Class TomlSyntaxException
        Inherits Exception

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace