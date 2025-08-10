Option Strict On
Option Explicit On

Imports ZoppaTomlLibrary.Collections
Imports ZoppaTomlLibrary.Strings

Namespace Toml

    ''' <summary>
    ''' TOMLノードを表すクラス。
    ''' </summary>
    ''' <remarks>
    ''' このクラスは、TOMLのノードを表現し、ノードの操作を提供します。
    ''' ノードは、値やテーブルなどのエントリーを持つことができます。
    ''' </remarks>
    Public Class TomlNode
        Inherits TomlElement

        ' エントリーツリー
        Protected ReadOnly _entryTree As Btree(Of EntryBase)

        ''' <summary>新しいTomlNodeを初期化します。</summary>
        ''' <param name="nodeType">ノードの種類。</param>
        ''' <remarks>このコンストラクタは、ノードの種類を指定して新しいインスタンスを初期化します。</remarks>
        Protected Sub New(nodeType As TomlNodeType)
            MyBase.New(nodeType)
            Me._entryTree = New Btree(Of EntryBase)(4)
        End Sub

        ''' <summary>
        ''' キーを指定して、値を登録します。
        ''' 
        ''' このメソッドは、キーと値のペアをエントリーツリーに登録します。
        ''' 同じキーが存在する場合、例外をスローします。
        ''' </summary>
        ''' <typeparam name="T">エントリーの型。</typeparam>
        ''' <param name="keyStr">キー文字列。</param>
        ''' <param name="entry">登録するエントリー。</param>
        ''' <remarks>このメソッドは、エントリーの型がEntryBaseを継承している必要があります。</remarks>
        Public Sub RegisterEntry(Of T As EntryBase)(keyStr As U8String, entry As T)
            If Me._entryTree.Contains(entry) Then
                ' 既に同じキーが存在する場合、例外をスロー
                Throw New TomlKeyDuplicationException($"キー '{keyStr}' はすでに存在します。")
            End If
            Me._entryTree.Insert(entry)
        End Sub

        ''' <summary>キーをドットで区切って、値を登録します。</summary>
        ''' <param name="keys">キーのリスト。</param>
        ''' <param name="index">現在のインデックス。</param>
        ''' <param name="tomlExpression">登録する値の式。</param>
        ''' <remarks>中間キーはテーブルノードとして登録されます。</remarks>
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
                Dim nextNode As TomlNode
                Dim entry = Me.GetEntry(curKey)
                Select Case entry?.Type
                    Case EntryType.Node
                        Select Case CType(entry, NodeEntry).Node.NodeType
                            Case TomlNodeType.Table ' 1
                                ' 中間キーがテーブルノードの場合、問題なし
                                nextNode = CType(CType(entry, NodeEntry).Node, TomlNode)
                            Case TomlNodeType.InlineTable   ' 2
                                Throw New TomlTableDuplicationException($"インラインテーブル'{curKey}'にキーを追加することはできません。")
                            Case Else   ' 3
                                Throw New TomlSyntaxException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                        End Select

                    Case EntryType.Value    ' 4
                        Throw New TomlSyntaxException($"キー '{curKey}' は値として登録されているため、テーブルノードとして登録できません。")

                    Case Else   ' 5
                        nextNode = New TomlTable(Me.IsDescribed)
                        Me.RegisterEntry(curKey, New NodeEntry(curKey, nextNode))
                End Select

                ' 次のキーが存在する場合、再帰的に登録
                nextNode.RegisterKeyValuePair(keys, index + 1, tomlExpression)
            Else
                ' 最後のキーの場合、値を登録
                Select Case tomlExpression.Type
                    Case TomlExpressionType.InlineTable
                        ' インラインテーブルの場合、ノードツリーに登録
                        Me.RegisterEntry(curKey, New NodeEntry(curKey, CreateInlineTable(tomlExpression)))

                    Case TomlExpressionType.Array
                        ' 配列の場合、配列ノードを作成してノードツリーに登録
                        Me.RegisterEntry(curKey, New NodeEntry(curKey, CreateArray(tomlExpression)))

                    Case Else
                        ' 通常の値の場合、値ツリーに登録
                        Me.RegisterEntry(curKey, New ValueEntry(curKey, tomlExpression))
                End Select
            End If
        End Sub

        ''' <summary>キーをドットで区切って、テーブルノードを登録します。</summary>
        ''' <param name="keys">キー。</param>
        ''' <param name="index">キーのドットレベル。</param>
        ''' <returns>Tomlノード。</returns>
        Function RegisterTable(keys As List(Of U8String), index As Integer, isInline As Boolean) As TomlTable
            Dim curKey = keys(index)

            Dim nextNode As TomlNode
            Dim entry = Me.GetEntry(curKey)

            ' キーのテーブルを検索、なければ作成し、あればチェックする
            '
            ' 1. キーのエントリがテーブルで明示的に登録されていなければ、問題なし
            ' 2. キーのエントリがインラインテーブルならば、重複例外をスロー
            ' 3. キーのエントリがテーブル配列で明示的に登録されていなければ、問題なし（最新位置のテーブル取得）
            ' 4. キーのエントリがテーブルでない場合、例外をスロー
            ' 5. キーのエントリが値として登録しているならば、例外をスロー
            ' 6. キーが登録されていない場合、新しいテーブルノードを作成
            Select Case entry?.Type
                Case EntryType.Node
                    Dim tblNode = CType(entry, NodeEntry).Node
                    Select Case tblNode.NodeType
                        Case TomlNodeType.Table ' 1
                            If tblNode.IsDescribed AndAlso index = keys.Count - 1 Then
                                Throw New TomlTableDuplicationException($"テーブル'{curKey}' はすでに存在します。")
                            End If
                            nextNode = CType(tblNode, TomlNode)

                        Case TomlNodeType.InlineTable   ' 2
                            Throw New TomlTableDuplicationException($"テーブル'{curKey}' はすでに存在します。")

                        Case TomlNodeType.ArrayTable    ' 3
                            If index = keys.Count - 1 Then
                                Throw New TomlTableDuplicationException($"テーブル'{curKey}' はすでに配列として定義されています。")
                            End If
                            nextNode = CType(tblNode, TomlArrayTable).GetLastTable()

                        Case Else   ' 4
                            Throw New TomlSyntaxException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                    End Select

                Case EntryType.Value    ' 5
                    Throw New TomlSyntaxException($"キー '{curKey}' は値として登録されているため、テーブルノードとして登録できません。")

                Case Else   ' 6
                    Dim described = (index = keys.Count - 1)
                    If isInline Then
                        nextNode = New TomlInlineTable()
                    Else
                        nextNode = New TomlTable(described)
                    End If
                    Me.RegisterEntry(curKey, New NodeEntry(curKey, nextNode))
            End Select

            ' 次のキーが存在する場合、再帰的に登録
            If index < keys.Count - 1 Then
                Return nextNode.RegisterTable(keys, index + 1, isInline)
            Else
                Return TryCast(nextNode, TomlTable)
            End If
        End Function

        ''' <summary>キーをドットで区切って、テーブル配列を登録します。</summary>
        ''' <param name="keys">キーのリスト。</param>
        ''' <param name="index">現在のインデックス。</param>
        ''' <returns>登録されたテーブル配列。</returns>
        ''' <remarks>中間キーはテーブルノードとして登録されます。</remarks>
        Function RegisterArrayTable(keys As List(Of U8String), index As Integer) As TomlArrayTable
            Dim curKey = keys(index)
            Dim entry = Me.GetEntry(keys(index))

            ' キーのテーブルを検索、なければ作成し、あればチェックする
            '
            ' 1. 中間キーのエントリがテーブルで明示的に登録されていなければ、問題なし
            ' 2. 中間キーのエントリがインラインテーブルならば、重複例外をスロー
            ' 3. 中間キーのエントリがテーブル配列で明示的に登録されていなければ、問題なし（最新位置のテーブル取得）
            ' 4. 末端キーのエントリがテーブル配列ならば新しいテーブルを追加
            ' 5. 末端キーのエントリがテーブルでない場合、例外をスロー
            ' 6. キーのエントリが値として登録しているならば、例外をスロー
            ' 7. 中間キーが登録されていない場合、新しいテーブルを作成
            ' 8. 末端キーが登録されていない場合、新しいテーブル配列を作成
            Select Case entry?.Type
                Case EntryType.Node
                    Dim tblNode = CType(entry, NodeEntry).Node
                    If index < keys.Count - 1 Then
                        Select Case tblNode.NodeType
                            Case TomlNodeType.Table ' 1
                                If Me.IsDescribed Then
                                    Throw New TomlTableDuplicationException($"テーブル'{curKey}' はすでに存在します。")
                                End If
                                Return CType(tblNode, TomlNode).RegisterArrayTable(keys, index + 1)
                            Case TomlNodeType.InlineTable
                                Throw New TomlTableDuplicationException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                            Case TomlNodeType.ArrayTable    ' 3
                                Dim midTable = CType(tblNode, TomlArrayTable).GetLastTable()
                                Return midTable.RegisterArrayTable(keys, index + 1)
                            Case Else
                                Throw New TomlSyntaxException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                        End Select
                    Else
                        Select Case tblNode.NodeType
                            Case TomlNodeType.ArrayTable    ' 4
                                Dim resTable = TryCast(tblNode, TomlArrayTable)
                                resTable.AppendNewTable()
                                Return resTable
                            Case Else   ' 5
                                Throw New TomlSyntaxException($"キー'{curKey}'はテーブル配列でなければなりません。")
                        End Select
                    End If

                Case EntryType.Value    ' 6
                    Throw New TomlSyntaxException($"キー '{curKey}' は値として登録されているため、テーブルノードとして登録できません。")

                Case Else
                    If index < keys.Count - 1 Then  ' 7
                        Dim rootTable = New TomlTable(False)
                        Me.RegisterEntry(curKey, New NodeEntry(curKey, rootTable))
                        Return rootTable.RegisterArrayTable(keys, index + 1)
                    Else
                        Dim resTable = New TomlArrayTable() ' 8
                        Me.RegisterEntry(curKey, New NodeEntry(curKey, resTable))
                        resTable.AppendNewTable()
                        Return resTable
                    End If
            End Select
        End Function

#Region "式、テーブルを取得"

        ''' <summary>
        ''' キーを指定して、値を取得します。
        ''' 
        ''' このメソッドは、指定されたキーに対応するエントリーを取得します。
        ''' キーが存在しない場合は、Nothingを返します。
        ''' </summary>
        ''' <param name="keyStr">キー文字列。</param>
        ''' <returns>対応するエントリー。</returns>
        ''' <remarks>このメソッドは、EntryBase型のエントリーを返します。</remarks>
        Function GetEntry(keyStr As U8String) As EntryBase
            Return Me._entryTree.Search(New EntryBase(keyStr))
        End Function

        ''' <summary>キーに対応するエントリーを再帰的に取得します。</summary>
        ''' <param name="keys">キーのリスト。</param>
        ''' <param name="index">現在のインデックス。</param>
        ''' <returns>対応するエントリー。</returns>
        Private Function GetRecursiveEntry(keys As List(Of U8String), index As Integer) As EntryBase
            Dim curKey = keys(index)
            If index < keys.Count - 1 Then
                ' 中間キーの場合、再帰的に取得
                Dim nextNode = TryCast(Me.GetEntry(curKey), NodeEntry)?.Node
                If nextNode Is Nothing Then
                    Throw New KeyNotFoundException($"キー'{curKey}'が見つかりません。")
                ElseIf nextNode.NodeType <> TomlNodeType.Table AndAlso nextNode.NodeType <> TomlNodeType.InlineTable Then
                    Throw New InvalidOperationException($"中間キー'{curKey}'はテーブルノードでなければなりません。")
                End If
                Return CType(nextNode, TomlNode).GetRecursiveEntry(keys, index + 1)
            Else
                ' 最後のキーの場合、値エントリーを取得
                Return Me.GetEntry(curKey)
            End If
        End Function

        ''' <summary>キーに対応するテーブルノードが存在するかどうかを確認します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>キーが存在する場合はTrue、存在しない場合はFalse。</returns>
        Public Overrides Function ContainsKey(key As String) As Boolean
            Dim res = Me.GetRecursiveEntry(AnalysisKey(key), 0)
            Return (res IsNot Nothing)
        End Function

        ''' <summary>キーをドットで区切って、UTF-8文字列のリストに変換します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>UTF-8文字列のリスト。</returns>
        Private Function AnalysisKey(key As String) As List(Of U8String)
            ' キーをUTF-8文字列に変換、式を解析
            Dim keyStr = U8String.NewString(key)
            Dim keys = TomlParser.ParseKey(keyStr, keyStr.GetIterator())

            ' キーを変換
            Return ConvertKeys(keys)
        End Function

        ''' <summary>キーに対応する式を取得し、存在しない場合は偽を返します。</summary>
        ''' <param name="key">キー。</param>
        ''' <param name="expression">取得した式。</param>
        ''' <returns>キーが存在する場合はTrue、存在しない場合はFalse。</returns>
        Public Function TryGetExpression(key As String, ByRef expression As TomlExpression) As Boolean
            Dim entry = Me.GetRecursiveEntry(AnalysisKey(key), 0)
            If entry.Type = EntryType.Value Then
                expression = CType(entry, ValueEntry).Value
                Return True
            Else
                expression = TomlExpression.Empty
                Return False
            End If
        End Function

        ''' <summary>キーに対応する式を取得します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>対応する式。</returns>
        ''' <remarks>キーが存在しない場合は、空の式を返します。</remarks>
        Public Overrides Function GetExpression(key As String) As TomlExpression
            Dim entry = Me.GetRecursiveEntry(AnalysisKey(key), 0)
            Return If(entry.Type = EntryType.Value, CType(entry, ValueEntry).Value, TomlExpression.Empty)
        End Function

        ''' <summary>キーに対応するテーブルノードを取得し、存在しない場合は偽を返します。</summary>
        ''' <param name="key">キー。</param>
        ''' <param name="table">取得したテーブルノード。</param>
        ''' <returns>キーが存在する場合はTrue、存在しない場合はFalse。</returns>
        Public Function TryGetNode(key As String, ByRef table As TomlElement) As Boolean
            Dim entry = Me.GetRecursiveEntry(AnalysisKey(key), 0)
            If entry.Type = EntryType.Node Then
                table = CType(entry, NodeEntry).Node
                Return True
            Else
                table = TomlElement.Empty
                Return False
            End If
        End Function

        ''' <summary>キーに対応するテーブルノードを取得します。</summary>
        ''' <param name="key">キー。</param>
        ''' <returns>対応するテーブルノード。</returns>
        Public Overrides Function GetNode(key As String) As TomlElement
            Dim entry = Me.GetRecursiveEntry(AnalysisKey(key), 0)
            Select Case entry.Type
                Case EntryType.Node
                    ' ノードエントリーの場合、ノードを返す
                    Return CType(entry, NodeEntry).Node

                Case EntryType.Value
                    ' 値エントリーの場合、値を持つノードを返す
                    Return New TomlValue(CType(entry, ValueEntry).Value)

                Case Else
                    ' エントリーが存在しない場合、例外
                    Throw New KeyNotFoundException($"キー '{key}' が見つかりません。")
            End Select
        End Function

        ''' <summary>
        ''' 登録されているキーのリストを取得します。
        ''' 
        ''' このメソッドは、現在の要素に登録されているキーのリストを返します。
        ''' 例えば、テーブルノードや配列テーブルノードで使用されます。
        ''' </summary>
        ''' <returns>キーのリスト。</returns>
        Public Overrides Function GetKeys() As IEnumerable(Of String)
            Return Me._entryTree.Select(Function(entry) entry.Key.ToString())
        End Function

#End Region

#Region "Entry"

        ''' <summary>エントリーの種類を表す列挙型。</summary>
        ''' <remarks>この列挙型は、TOMLエントリーの種類を定義します。</remarks>
        Enum EntryType
            ''' <summary>エントリーの種類が指定されていない。</summary>
            None

            ''' <summary>キーと値のペアを表すエントリー。</summary>
            Value

            ''' <summary>ノードを表すエントリー。</summary>
            Node
        End Enum

        ''' <summary>エントリーの基底クラス。</summary>
        ''' <remarks>このクラスは、キーと値のペアやノードを表すエントリーの基底クラスです。</remarks>
        Class EntryBase
            Implements IComparable(Of EntryBase)

            ''' <summary>エントリーの種類。</summary>
            ''' <remarks>このプロパティは、エントリーの種類を示します。</remarks>
            Public ReadOnly Property Type As EntryType

            ''' <summary>エントリーのキー。</summary>
            ''' <remarks>このプロパティは、エントリーのキーを示します。</remarks>
            Public ReadOnly Property Key As U8String

            ''' <summary>新しいエントリーを初期化します。</summary>
            ''' <param name="key">エントリーのキー。</param>
            Public Sub New(key As U8String)
                Me.Type = EntryType.None
                Me.Key = key
            End Sub

            ''' <summary>新しいエントリーを初期化します。</summary>
            ''' <param name="etype">エントリーの種類。</param>
            ''' <param name="key">エントリーのキー。</param>
            Public Sub New(etype As EntryType, key As U8String)
                Me.Type = etype
                Me.Key = key
            End Sub

            ''' <summary>キーの比較を行います。</summary>
            ''' <param name="other">比較対象。</param>
            ''' <returns>比較結果。</returns>
            Public Function CompareTo(other As EntryBase) As Integer Implements IComparable(Of EntryBase).CompareTo
                Return Me.Key.CompareTo(other.Key)
            End Function

        End Class

        ''' <summary>値エントリーを表すクラス。</summary>
        ''' <remarks>このクラスは、キーと値のペアを表すエントリーです。</remarks>
        NotInheritable Class ValueEntry
            Inherits EntryBase

            ''' <summary>値を表す式を取得します。</summary>
            Public ReadOnly Property Value As TomlExpression

            ''' <summary>新しい値エントリーを初期化します。</summary>
            ''' <param name="key">エントリーのキー。</param>
            ''' <param name="value">値を表す式。</param>
            Public Sub New(key As U8String, value As TomlExpression)
                MyBase.New(EntryType.Value, key)
                Me.Value = value
            End Sub

            ''' <summary>値エントリーの文字列表現を返します。</summary>
            ''' <returns>値エントリーの文字列表現。</returns>
            Overrides Function ToString() As String
                Return $"{Key} : {Value.Type} = {Value.Str}"
            End Function

        End Class

        ''' <summary>ノードエントリーを表すクラス。</summary>
        ''' <remarks>このクラスは、キーとノードを表すエントリーです。</remarks>
        NotInheritable Class NodeEntry
            Inherits EntryBase

            ''' <summary>子ノードを取得します。</summary>
            Public Property Node As TomlElement

            '''' <summary>新しいノードエントリーを初期化します。</summary>
            ''' <param name="key">エントリーのキー。</param>
            ''' <param name="node">子ノード。</param>
            Public Sub New(key As U8String, node As TomlElement)
                MyBase.New(EntryType.Node, key)
                Me.Node = node
            End Sub

            ''' <summary>ノードエントリーの文字列表現を返します。</summary>
            ''' <returns>ノードエントリーの文字列表現。</returns>
            Overrides Function ToString() As String
                Return $"{Key} : {Node.NodeType}"
            End Function

        End Class

#End Region

    End Class

End Namespace
