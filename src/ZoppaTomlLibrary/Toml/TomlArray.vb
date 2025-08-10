Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>
    ''' TOMLの配列を表すクラス。
    ''' このクラスは、TOMLの配列を表現し、配列の操作を提供します。
    ''' </summary>
    ''' <remarks>
    ''' TOMLの配列は、複数の値を格納するためのコンテナです。
    ''' 各値は、TOMLの要素として表現されます。
    ''' </remarks>
    Public NotInheritable Class TomlArray
        Inherits TomlElement

        ' 配列テーブルのリスト
        Private ReadOnly _items As New List(Of TomlElement)

        ''' <summary>要素の数を取得します。</summary>
        ''' <remarks>このプロパティは、要素の数を返します。デフォルトでは1を返します。</remarks>
        ''' <returns>要素の数。</returns>
        Public Overrides ReadOnly Property Count As Integer
            Get
                Return Me._items.Count
            End Get
        End Property

        ''' <summary>インデックスに対応する要素を取得します。</summary>
        ''' <param name="index">インデックス。</param>
        ''' <returns>対応する要素。</returns>
        ''' <remarks>デフォルトでは、空の式を返します。</remarks>
        Default Public Overrides ReadOnly Property Item(index As Integer) As TomlElement
            Get
                If index < 0 OrElse index >= Me._items.Count Then
                    Throw New IndexOutOfRangeException($"インデックスの値が範囲を超えています：{index}")
                End If
                Return Me._items(index)
            End Get
        End Property

        ''' <summary>
        ''' 新しいTOML配列を初期化します。
        ''' このコンストラクタは、配列のノードタイプを設定します。
        ''' </summary>
        Public Sub New()
            MyBase.New(TomlNodeType.Array)
        End Sub

        ''' <summary>
        ''' TOMLの配列に新しい要素を追加します。
        ''' このメソッドは、配列に新しいTOML要素を追加します。
        ''' </summary>
        ''' <param name="itemNode">TOML要素。</param>
        Sub AddItem(itemNode As TomlElement)
            Me._items.Add(itemNode)
        End Sub

        ''' <summary>
        ''' TOMLの配列に新しい要素を追加します。
        ''' このメソッドは、配列に新しいTOML式を追加します。
        ''' </summary>
        ''' <param name="itemExpr">TOML式。</param>
        ''' <remarks>このメソッドは、TOMLの値を表すノードを作成して追加します。</remarks>
        Sub AddItem(itemExpr As TomlExpression)
            Me._items.Add(New TomlValue(itemExpr))
        End Sub

    End Class

End Namespace
