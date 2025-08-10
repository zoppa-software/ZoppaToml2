Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLのインラインテーブルを表すクラス。
    ''' このクラスは、TOMLのインラインテーブルを表現し、インラインテーブルの操作を提供します。
    ''' </summary>
    ''' <remarks>
    ''' インラインテーブルは、通常の値ツリーではなく、ノードツリーに登録されます。
    ''' </remarks>
    Public NotInheritable Class TomlInlineTable
        Inherits TomlNode

        ''' <summary>コンストラクタ。</summary>
        ''' <remarks>
        ''' このコンストラクタは、TOMLのインラインテーブルを表すノードを作成します。
        ''' </remarks>
        Public Sub New()
            MyBase.New(TomlNodeType.InlineTable)
        End Sub

    End Class

End Namespace
