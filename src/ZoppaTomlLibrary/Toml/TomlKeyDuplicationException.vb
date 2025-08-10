Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>
    ''' TOMLのキーの重複に関する例外を表すクラス。
    ''' このクラスは、TOMLの解析中に同じキーが複数回定義された場合にスローされます。
    ''' </summary>
    ''' <remarks>
    ''' この例外は、TOMLの解析中に同じキーが複数回定義された場合にスローされます。
    ''' 例えば、同じセクション内で同じキーが2回定義された場合などです。
    ''' </remarks>
    Public NotInheritable Class TomlKeyDuplicationException
        Inherits Exception

        ''' <summary>エラーの原因となったTOML文字列を取得します。</summary>
        Public ReadOnly Property ErrorSource As U8String

        ''' <summary>エラーの原因となったTOML文字列のイテレータを取得します。</summary>
        Public ReadOnly Property ErrorIter As U8String.U8StringIterator

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        Public Sub New(message As String)
            MyBase.New(message)
            Me.ErrorSource = U8String.Empty
            Me.ErrorIter = U8String.Empty.GetIterator()
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="message">例外メッセージ。</param>
        ''' <param name="source">解析対象のTOML文字列。</param>
        ''' <param name="iter">TOML文字列のイテレータ。</param>
        ''' <remarks>
        ''' このコンストラクタは、TOMLの解析中に発生したエラーを示すために使用されます。
        ''' </remarks>
        Public Sub New(message As String, source As U8String, iter As U8String.U8StringIterator)
            MyBase.New(message)
            Me.ErrorSource = source
            Me.ErrorIter = iter
        End Sub

    End Class

End Namespace