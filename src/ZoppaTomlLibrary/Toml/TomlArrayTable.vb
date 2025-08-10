Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Strings
Imports ZoppaTomlLibrary.Toml.TomlNode

Namespace Toml

    ''' <summary>
    ''' TOMLの配列テーブルを表すクラス。
    ''' このクラスは、TOMLの配列テーブルを表現し、配列テーブルの操作を提供します。
    ''' </summary>
    ''' <remarks>
    ''' 配列テーブルは、複数のTOMLテーブルを格納するためのコンテナです。
    ''' 各テーブルは、キーと値のペアを持つことができます。
    ''' </remarks>
    Public NotInheritable Class TomlArrayTable
        Inherits TomlElement

        ' 配列テーブルのリスト
        Private ReadOnly _tables As New List(Of TomlTable)

        ''' <summary>配列テーブルの数を取得します。</summary>
        ''' <returns>配列テーブルの数。</returns>
        ''' <remarks>このプロパティは、配列テーブルの数を返します。</remarks>
        Public Overrides ReadOnly Property Count As Integer
            Get
                Return _tables.Count
            End Get
        End Property

        ''' <summary>配列テーブルのインデックスを指定して、対応するTOMLテーブルを取得します。</summary>
        ''' <param name="index">インデックス。</param>
        ''' <returns>指定されたインデックスのTOMLテーブル。</returns>
        ''' <exception cref="ArgumentOutOfRangeException">インデックスが範囲外の場合にスローされます。</exception>
        Default Public Overrides ReadOnly Property Item(index As Integer) As TomlElement
            Get
                If index < 0 OrElse index >= Me._tables.Count Then
                    Throw New IndexOutOfRangeException($"インデックスの値が範囲を超えています：{index}")
                End If
                Return Me._tables(index)
            End Get
        End Property

        ''' <summary>新しいTOML配列テーブルを初期化します。</summary>
        ''' <remarks>このコンストラクタは、配列テーブルのノードタイプを設定します。</remarks>

        Public Sub New()
            MyBase.New(TomlNodeType.ArrayTable)
            Me._tables = New List(Of TomlTable)()
        End Sub

        ''' <summary>新しいTOMLテーブルを配列に追加します。</summary>
        ''' <remarks>このメソッドは、配列テーブルに新しいTOMLテーブルを追加します。</remarks>
        Public Sub AppendNewTable()
            Dim newTable As New TomlTable(False)
            Me._tables.Add(newTable)
        End Sub

        ''' <summary>
        ''' 値、ノードのエントリをテーブル配列に登録します。
        ''' このメソッドは、指定されたキー文字列とエントリを使用して、末尾のTOMLテーブルにエントリを登録します。
        ''' </summary>
        ''' <param name="keyStr">キー。</param>
        ''' <param name="entry">エントリ。</param>
        Public Sub RegisterEntryToArray(Of T As EntryBase)(keyStr As U8String, entry As T)
            Me._tables.Last().RegisterEntry(keyStr, entry)
        End Sub

        ''' <summary>最後のTOMLテーブルを取得します。</summary>
        ''' <returns>最後のTOMLテーブル。</returns>
        Public Function GetLastTable() As TomlTable
            Return Me._tables.Last()
        End Function

        ''' <summary>キーをドットで区切って、値をテーブル配列に登録します。</summary>
        ''' <param name="keys">キーのリスト。</param>
        ''' <param name="index">現在のインデックス。</param>
        ''' <param name="tomlExpression">登録する値の式。</param>
        Public Overrides Sub RegisterKeyValuePair(keys As List(Of U8String), index As Integer, tomlExpression As TomlExpression)
            Dim curKey = keys(index)
            If index < keys.Count - 1 Then
                ' 中間キーの場合、再帰的に登録
                '
                ' 1. キーのエントリがテーブルならば問題なし
                ' 2. キーのエントリがインラインテーブルならば、重複例外をスロー
                ' 3. キーのエントリがテーブルでない場合、例外をスロー
                ' 4. キーのエントリが値として登録しているならば、例外をスロー
                ' 5. キーが登録されていない場合、新しいテーブルノードを作成
                Dim nextNode As TomlElement
                Dim entry = Me.GetLastTable().GetEntry(curKey)
                Select Case entry?.Type
                    Case EntryType.Node
                        nextNode = CType(entry, NodeEntry).Node
                        Select Case nextNode.NodeType
                            Case TomlNodeType.Table ' 1
                                ' 中間キーがテーブルノードの場合、問題なし
                            Case TomlNodeType.InlineTable   ' 2
                                Throw New TomlTableDuplicationException($"インラインテーブル'{curKey}'にキーを追加することはできません。")
                            Case Else   ' 3
                                Throw New TomlSyntaxException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                        End Select

                    Case EntryType.Value    ' 4
                        Throw New TomlSyntaxException($"キー '{curKey}' は値として登録されているため、テーブルノードとして登録できません。")

                    Case Else   ' 5
                        nextNode = New TomlTable(Me.IsDescribed)
                        Me.GetLastTable().RegisterEntry(curKey, New NodeEntry(curKey, nextNode))
                End Select

                ' 次のキーが存在する場合、再帰的に登録
                nextNode.RegisterKeyValuePair(keys, index + 1, tomlExpression)
            Else
                ' 最後のキーの場合、値を登録
                Dim resNode = TryCast(Me, TomlArrayTable)
                Select Case tomlExpression.Type
                    Case TomlExpressionType.InlineTable
                        ' インラインテーブルの場合、ノードツリーに登録
                        resNode.RegisterEntryToArray(curKey, New NodeEntry(curKey, CreateInlineTable(tomlExpression)))

                    Case Else
                        ' 通常の値の場合、値ツリーに登録
                        resNode.RegisterEntryToArray(curKey, New ValueEntry(curKey, tomlExpression))
                End Select
            End If
        End Sub

    End Class

End Namespace

