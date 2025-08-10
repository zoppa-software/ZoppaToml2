Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLのテーブルの重複に関する例外を表すクラス。
    ''' このクラスは、TOMLの解析中に同じテーブルが複数回定義された場合にスローされます。
    ''' </summary>
    ''' <remarks>
    ''' この例外は、TOMLの解析中に同じテーブルが複数回定義された場合にスローされます。
    ''' 例えば、同じセクション内で同じテーブルが2回定義された場合などです。
    ''' </remarks>
    Public NotInheritable Class TomlTableDuplicationException
        Inherits Exception

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

    End Class

End Namespace
