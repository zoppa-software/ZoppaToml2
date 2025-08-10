Option Strict On
Option Explicit On

Namespace Toml

    ''' <summary>TOMLノードの種類を表す列挙型。</summary>
    ''' <remarks>
    ''' この列挙型は、TOMLのノードの種類を定義します。
    ''' 各ノードは、特定の種類に分類され、TOMLの構造を表現します。
    ''' </remarks>
    Public Enum TomlNodeType

        ''' <summary>ノードの種類が指定されていない。</summary>
        None

        ''' <summary>値を表す。</summary>
        Value

        ''' <summary>キーと値のペアの集合。</summary>
        Keyvals

        ''' <summary>テーブル（セクション）を表す。</summary>
        Table

        ''' <summary>インラインテーブルを表す。</summary>
        InlineTable

        ''' <summary>テーブル配列を表す。</summary>
        ArrayTable

        ''' <summary>配列を表す。</summary>
        Array

    End Enum

End Namespace
