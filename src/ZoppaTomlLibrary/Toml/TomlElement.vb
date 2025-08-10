Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>TOMLの要素を表すクラスです。</summary>
    ''' <remarks>
    ''' このクラスは、TOMLの要素を表す基本クラスであり、ノードの種類を保持します。
    ''' </remarks>
    Public Class TomlElement

        ' 空のインスタンスを生成するためのLazy初期化
        Private Shared ReadOnly _empty As New Lazy(Of TomlElement)(Function() New TomlElement(TomlNodeType.None))

        ''' <summary>ノードの種類を表します。</summary>
        Public ReadOnly Property NodeType As TomlNodeType

        ''' <summary>記述されているテーブルならばTrueを返します。</summary>
        Public Overridable ReadOnly Property IsDescribed As Boolean = False

        ''' <summary>空のノードを取得します。</summary>
        ''' <remarks>このプロパティは、空のノードを返します。</remarks>
        ''' <returns>空のノード。</returns>
        Public Shared ReadOnly Property Empty As TomlElement
            Get
                Return _empty.Value
            End Get
        End Property

        ''' <summary>要素の数を取得します。</summary>
        ''' <remarks>このプロパティは、要素の数を返します。デフォルトでは1を返します。</remarks>
        ''' <returns>要素の数。</returns>
        Public Overridable ReadOnly Property Count As Integer
            Get
                Return 1
            End Get
        End Property

        ''' <summary>インデックスに対応する要素を取得します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>対応する要素。</returns>
        Default Public ReadOnly Property Item(key As String) As TomlElement
            Get
                Return Me.GetNode(key)
            End Get
        End Property

        ''' <summary>インデックスに対応する要素を取得します。</summary>
        ''' <param name="index">インデックス。</param>
        ''' <returns>対応する要素。</returns>
        ''' <remarks>デフォルトでは、空の式を返します。</remarks>
        Default Public Overridable ReadOnly Property Item(index As Integer) As TomlElement
            Get
                Return TomlElement.Empty
            End Get
        End Property

        ''' <summary>新しいTomlNodeを初期化します。</summary>
        ''' <param name="nodeType">ノードの種類。</param>
        ''' <remarks>このコンストラクタは、ノードの種類を指定して新しいインスタンスを初期化します。</remarks>
        Protected Sub New(nodeType As TomlNodeType)
            Me.NodeType = nodeType
        End Sub

        ''' <summary>式を取得します。</summary>
        ''' <returns>式。</returns>
        ''' <remarks>存在しない場合は、空の式を返します。</remarks>
        Public Overridable Function GetExpression() As TomlExpression
            Return TomlExpression.Empty
        End Function

        ''' <summary>キーに対応する要素が存在するかどうかを確認します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>キーが存在する場合はTrue、存在しない場合はFalse。</returns>
        Public Overridable Function ContainsKey(key As String) As Boolean
            Return False
        End Function

        ''' <summary>キーに対応する式を取得します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>対応する式。</returns>
        ''' <remarks>キーが存在しない場合は、空の式を返します。</remarks>
        Public Overridable Function GetExpression(key As String) As TomlExpression
            Return TomlExpression.Empty
        End Function

        ''' <summary>キーに対応するテーブルノードを取得します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>対応するテーブルノード。</returns>
        ''' <remarks>キーが存在しない場合は、空のテーブルノードを返します。</remarks>
        Public Overridable Function GetNode(key As String) As TomlElement
            Return Me
        End Function

        ''' <summary>
        ''' 式の内容を指定した型に変換します。
        ''' このメソッドは、式の内容を指定された型に変換します。
        ''' 例えば、基本文字列やマルチライン基本文字列に変換する場合に使用されます。
        ''' </summary>
        ''' <typeparam name="T">型。</typeparam>
        ''' <returns>値。</returns>
        Public Overridable Function ValueTo(Of T)() As T
            Throw New NotSupportedException("式を持たないため値を取得できません")
        End Function

        ''' <summary>キーをドットで区切って、値を登録します。</summary>
        ''' <param name="keys">キーのリスト。</param>
        ''' <param name="index">現在のインデックス。</param>
        ''' <param name="tomlExpression">登録する値の式。</param>
        ''' <remarks>中間キーはテーブルノードとして登録されます。</remarks>
        Public Overridable Sub RegisterKeyValuePair(keys As List(Of U8String), index As Integer, tomlExpression As TomlExpression)
            Throw New NotSupportedException("テーブルノードではないため、キーと値のペアを登録できません。")
        End Sub

        ''' <summary>
        ''' 登録されているキーのリストを取得します。
        ''' 
        ''' このメソッドは、現在の要素に登録されているキーのリストを返します。
        ''' 例えば、テーブルノードや配列テーブルノードで使用されます。
        ''' </summary>
        ''' <returns>キーのリスト。</returns>
        Public Overridable Function GetKeys() As IEnumerable(Of String)
            Return New String() {}
        End Function

    End Class

End Namespace