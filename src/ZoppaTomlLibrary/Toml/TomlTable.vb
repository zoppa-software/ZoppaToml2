Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLのテーブルを表すクラス。
    ''' このクラスは、TOMLのテーブル（セクション）を表現します。
    ''' </summary>
    ''' <remarks>
    ''' TOMLのテーブルは、キーと値のペアの集合であり、セクションを表すために使用されます。
    ''' このクラスは、TOMLのテーブルを表すノードを作成します。
    ''' </remarks>
    Public NotInheritable Class TomlTable
        Inherits TomlNode

        ''' <summary>テーブルが記述されている場合にTrueを返します。</summary>
        Public Overrides ReadOnly Property IsDescribed As Boolean

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="described">テーブルが記述されているかどうか。</param>
        ''' <remarks>
        ''' このコンストラクタは、TOMLのテーブルを表すノードを作成します。
        ''' </remarks>
        Public Sub New(described As Boolean)
            MyBase.New(TomlNodeType.Table)
            Me.IsDescribed = described
        End Sub

    End Class

End Namespace

