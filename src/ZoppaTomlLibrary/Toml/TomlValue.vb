Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLの値を表すクラス。
    ''' このクラスは、TOMLの値を表現し、値の操作を提供します。
    ''' </summary>
    ''' <remarks>
    ''' TOMLの値は、文字列、数値、真偽値などの基本的なデータ型を表します。
    ''' このクラスは、TOMLの値を表すノードを作成します。
    ''' </remarks>
    Public NotInheritable Class TomlValue
        Inherits TomlElement

        ' 対象の式
        Private ReadOnly _element As TomlExpression

        ''' <summary>新しいArrayEntryを初期化します。</summary>
        ''' <param name="element">配列の要素。</param>
        Public Sub New(element As TomlExpression)
            MyBase.New(TomlNodeType.Value)
            Me._element = element
        End Sub

        ''' <summary>式を取得します。</summary>
        ''' <returns>式。</returns>
        ''' <remarks>存在しない場合は、空の式を返します。</remarks>
        Public Overrides Function GetExpression() As TomlExpression
            Return Me._element
        End Function

        ''' <summary>
        ''' 式の内容を指定した型に変換します。
        ''' このメソッドは、式の内容を指定された型に変換します。
        ''' 例えば、基本文字列やマルチライン基本文字列に変換する場合に使用されます。
        ''' </summary>
        ''' <typeparam name="T">型。</typeparam>
        ''' <returns>値。</returns>
        Public Overrides Function ValueTo(Of T)() As T
            Return Me._element.ValueTo(Of T)()
        End Function

    End Class

End Namespace