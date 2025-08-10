Option Strict On
Option Explicit On

Namespace Collections

    ''' <summary>
    ''' B木の実装クラス
    ''' 
    ''' B木は、データベースやファイルシステムなどで使用される自己平衡二分探索木の一種です。
    ''' B木は、ノードが複数のキーと子ノードを持つことができ、効率的な検索、挿入、および削除操作を提供します。
    ''' 
    ''' このクラスは、ジェネリック型パラメータ T を使用して、任意の比較可能な型の要素を格納できます。
    ''' </summary>
    ''' <typeparam name="T">対象の型。</typeparam>
    Public NotInheritable Class Btree(Of T As IComparable(Of T))
        Implements IEnumerable(Of T)

        ' B木の要素数
        Private ReadOnly M As Integer = 4

        ' B木のルートノード
        Private _root As Node

        ' 格納されているデータ数
        Private _count As Integer

        ' ロックオブジェクト
        Private ReadOnly _lock As New Object()

        ''' <summary>
        ''' B木のデータ数を取得するプロパティ。
        ''' このプロパティは、B木に格納されている要素の数を返します。
        ''' </summary>
        ''' <returns>データ数。</returns>
        Private ReadOnly Property Count As Integer
            Get
                SyncLock Me._lock
                    Return Me._count
                End SyncLock
            End Get
        End Property

        ''' <summary>コンストラクタ。</summary>
        Public Sub New()
            Me.M = 4
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="blockSize">ブロックサイズ指定</param>
        Public Sub New(blockSize As UShort)
            Me.M = blockSize \ 2
        End Sub

        ''' <summary>
        ''' B木に要素を挿入するメソッド。
        ''' 
        ''' このメソッドは、指定された値をB木に挿入します。
        ''' 挿入操作は、B木の特性を維持するために行われます。
        ''' 
        ''' 挿入後、必要に応じてノードが分割され、新しいルートノードが作成されることがあります。
        ''' </summary>
        ''' <param name="value">挿入する値。</param>
        Public Sub Insert(value As T)
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value), "挿入する値はnullにできません")
            End If

            ' 挿入操作の実装
            Dim update As New UpdateInfo With {
                .Value = value,
                .Done = False,
                .NewNode = Nothing
            }

            SyncLock Me._lock
                ' ルートノードに挿入を試みる
                If Not InsertIntoNode(Me._root, update).Done Then
                    Dim newNode = New Node(Me.M)
                    newNode.values(0) = update.Value
                    newNode.Count = 1
                    newNode.children(0) = Me._root
                    newNode.children(1) = update.NewNode
                    newNode.Count = 1

                    ' ルートノードを更新する
                    Me._root = newNode
                End If

                ' 数を更新
                Me._count += 1
            End SyncLock
        End Sub

        ''' <summary>
        ''' B木のノードに値を挿入するヘルパーメソッド。
        ''' このメソッドは、指定されたノードに値を挿入し、必要に応じてノードを分割します。
        ''' 挿入が成功した場合は、更新情報を返します。
        ''' </summary>
        ''' <param name="node">挿入先のノード。</param>
        ''' <param name="update">更新情報。</param>
        ''' <returns>更新情報。</returns>
        Private Function InsertIntoNode(node As Node, update As UpdateInfo) As UpdateInfo
            If node IsNot Nothing Then
                ' 指定値が挿入される位置を取得
                Dim i As Integer = 0
                While i < node.Count AndAlso node.values(i).CompareTo(update.Value) < 0
                    i += 1
                End While

                ' 既に同じ値が存在する場合は、エラーを返す
                If i < node.Count AndAlso node.values(i).CompareTo(update.Value) = 0 Then
                    update.Done = True
                    Throw New BtreeException("既に値が登録されています")
                End If

                ' 登録されているノードが存在しない場合は、子のノードへ挿入
                If Not Me.InsertIntoNode(node.children(i), update).Done Then
                    If node.Count < node.values.Length Then
                        ' ノードを割らずに挿入できる場合は、値を挿入する
                        Me.InsertValueIntoNode(node, i, update)
                        update.Done = True
                    Else
                        ' ノードが満杯の場合、ノードを分割する
                        Me.SplitAndIntoNode(node, i, update)
                        update.Done = False
                    End If
                End If
            Else
                ' 挿入先のノードがない場合、呼び元のノードへ挿入する
                update.Done = False
                update.NewNode = Nothing
            End If
            Return update
        End Function

        ''' <summary>
        ''' ノードに値を挿入するメソッド。
        ''' このメソッドは、指定されたノードに値を挿入します。
        ''' 挿入後、子ノードの配列が更新されます。
        ''' </summary>
        ''' <param name="node">挿入先のノード。</param>
        ''' <param name="index">挿入位置のインデックス。</param>
        ''' <param name="update">更新情報。</param>
        Private Sub InsertValueIntoNode(node As Node, index As Integer, update As UpdateInfo)
            ' ノードの値を右にシフト
            For i As Integer = node.Count - 1 To index Step -1
                node.values(i + 1) = node.values(i)
                node.children(i + 2) = node.children(i + 1)
            Next
            node.values(index) = update.Value
            node.children(index + 1) = update.NewNode

            node.Count += 1
        End Sub

        ''' <summary>
        ''' ノードを分割して新しいノードに値を挿入するメソッド。
        ''' このメソッドは、指定されたノードを分割し、新しいノードに値を挿入します。
        ''' </summary>
        ''' <param name="node">分割対象のノード。</param>
        ''' <param name="index">挿入位置のインデックス。</param>
        ''' <param name="update">更新情報。</param>
        Private Sub SplitAndIntoNode(node As Node, index As Integer, update As UpdateInfo)
            Dim newNode = New Node(Me.M) With {
                .Count = 0
            }

            ' 分割位置を計算
            Dim splitIndex = If(index <= M, M, M + 1)

            ' 右側のノードに値をコピー
            Dim j As Integer = 0
            For i As Integer = splitIndex To node.Count - 1
                newNode.values(j) = node.values(i)
                newNode.children(j + 1) = node.children(i + 1)
                j += 1
            Next
            newNode.Count = node.values.Length - splitIndex

            ' 左側のノードは数のみ更新
            node.Count = splitIndex

            ' 値を挿入する
            If index <= M Then
                Me.InsertValueIntoNode(node, index, update)
            Else
                Me.InsertValueIntoNode(newNode, index - splitIndex, update)
            End If

            ' 呼び出し元で更新する情報を設定
            update.Value = node.values(node.Count - 1)
            update.NewNode = newNode

            ' リンクを更新
            newNode.children(0) = node.children(node.Count)
            node.Count -= 1
        End Sub

        ''' <summary>
        ''' B木に値が含まれているかどうかを確認するメソッド。
        ''' このメソッドは、指定された値がB木に存在するかどうかを確認します。
        ''' 存在する場合はTrueを返し、存在しない場合はFalseを返します。
        ''' </summary>
        ''' <param name="value">検索する値。</param>
        ''' <returns>値が含まれている場合はTrue、そうでない場合はFalse。</returns>
        Public Function Contains(value As T) As Boolean
            SyncLock Me._lock
                Dim ptr As Node = Me._root
                While ptr IsNot Nothing
                    Dim i As Integer = 0
                    While i < ptr.Count AndAlso ptr.values(i).CompareTo(value) < 0
                        i += 1
                    End While
                    If i < ptr.Count AndAlso ptr.values(i).CompareTo(value) = 0 Then
                        Return True ' 値が見つかった
                    End If
                    ' 次の子ノードへ移動
                    ptr = ptr.children(i)
                End While
                Return False ' 値が見つからなかった
            End SyncLock
        End Function

        ''' <summary>
        ''' B木から値を検索するメソッド。
        ''' このメソッドは、指定された値をB木から検索し、見つかった場合はその値を返します。
        ''' 見つからなかった場合は、Nothing を返します。
        ''' </summary>
        ''' <param name="value">検索する値。</param>
        ''' <returns>見つかった値、または Nothing。</returns>
        Public Function Search(value As T) As T
            SyncLock Me._lock
                Dim ptr As Node = Me._root
                While ptr IsNot Nothing
                    Dim i As Integer = 0
                    While i < ptr.Count AndAlso ptr.values(i).CompareTo(value) < 0
                        i += 1
                    End While
                    If i < ptr.Count AndAlso ptr.values(i).CompareTo(value) = 0 Then
                        Return ptr.values(i)
                    End If
                    ' 次の子ノードへ移動
                    ptr = ptr.children(i)
                End While
                Return CType(Nothing, T)
            End SyncLock
        End Function

        ''' <summary>
        ''' B木をクリアするメソッド。
        ''' このメソッドは、B木からすべての値を削除し、木を空にします。
        ''' ルートノードを Nothing に設定し、データ数を 0 にリセットします。
        ''' </summary>
        Public Sub Clear()
            SyncLock Me._lock
                Me._root = Nothing
                Me._count = 0
            End SyncLock
        End Sub

        ''' <summary>
        ''' B木から指定された値を削除するメソッド。
        ''' このメソッドは、B木から指定された値を削除し、必要に応じてノードを再構築します。
        ''' 削除後、ルートノードが空になった場合は、ルートを子要素で置き換えます。
        ''' </summary>
        ''' <param name="value">削除する値。</param>
        Public Sub Remove(value As T)
            If value Is Nothing Then
                Throw New ArgumentNullException(NameOf(value), "挿入する値はnullにできません")
            End If

            ' 削除操作の実装
            Dim delete As New DeleteInfo With {
                .Value = value,
                .Done = False,
                .NeedRebuild = False
            }
            SyncLock Me._lock
                ' ルートノードから削除を開始
                Me.RemoveFromNode(Me._root, delete)
                If delete.Done Then
                    If Me._root IsNot Nothing Then
                        ' ルートノードが空になった場合は、ルートを子要素で置き換える
                        If Me._root.Count = 0 Then
                            Me._root = Me._root.children(0)
                        End If
                    End If

                    ' 削除が成功した場合は、カウントを減らす
                    Me._count -= 1
                Else
                    ' 要素がなく、削除できなかった場合はエラーを返す
                    Throw New BtreeException("指定された値は存在せず、削除できません")
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' B木のノードから値を削除するヘルパーメソッド。
        ''' このメソッドは、指定されたノードから値を削除し、必要に応じてノードを再構築します。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub RemoveFromNode(curNode As Node, delete As DeleteInfo)
            If curNode IsNot Nothing Then
                ' 削除する値の位置を探す
                Dim del As Integer = 0
                While del < curNode.Count AndAlso curNode.values(del).CompareTo(delete.Value) < 0
                    del += 1
                End While

                ' 値が見つかった場合は削除する、そうでない場合は子ノードから削除を試みる
                If del < curNode.Count AndAlso curNode.values(del).CompareTo(delete.Value) = 0 Then
                    delete.Done = True
                    If curNode.children(del + 1) IsNot Nothing Then
                        Dim rnode = curNode.children(del + 1)

                        ' 右側の子ノードの最小値を取得して置き換える
                        Dim rem_node = rnode
                        While rem_node.children(0) IsNot Nothing
                            rem_node = rem_node.children(0)
                        End While

                        ' 右側の子ノードの最小値を削除
                        curNode.values(del) = rem_node.values(0)
                        delete.Value = rem_node.values(0)
                        Me.RemoveFromNode(curNode.children(del + 1), delete)
                        If delete.NeedRebuild Then
                            Me.RestoreNode(curNode, del + 1, delete)
                        End If
                    Else
                        Me.RemoveItem(curNode, del, delete)
                    End If
                Else
                    Me.RemoveFromNode(curNode.children(del), delete)
                    If delete.NeedRebuild Then
                        Me.RestoreNode(curNode, del, delete)
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' ノードから値を削除します。
        ''' 値が見つかった場合は、削除し、子ノードも削除します。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="del">削除された値のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub RemoveItem(curNode As Node, del As Integer, delete As DeleteInfo)
            ' 値を削除し、右側の値を左にシフト
            For i As Integer = del To curNode.Count - 2
                curNode.values(i) = curNode.values(i + 1)
                curNode.children(i + 1) = curNode.children(i + 2)
            Next

            ' ノードの値の数を減らす、小さくなりすぎたら needRebuild フラグを立てる
            curNode.Count -= 1
            delete.NeedRebuild = curNode.Count < M
        End Sub

        ''' <summary>
        ''' ノードを再構築するためのヘルパーメソッド。
        ''' このメソッドは、ノードが小さくなりすぎた場合に、隣接ノードから値を借りるか、結合します。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="del">削除された値のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub RestoreNode(curNode As Node, del As Integer, delete As DeleteInfo)
            delete.NeedRebuild = False
            If del > 0 Then
                If curNode.children(del - 1) IsNot Nothing Then
                    ' 左側のノードから値を借りる
                    Dim lNode = curNode.children(del - 1)
                    If lNode.Count > M Then
                        Me.MoveRight(curNode, del)
                    Else
                        Me.Combine(curNode, del, delete)
                    End If
                End If
            ElseIf curNode.children(1) IsNot Nothing Then
                ' 右側のノードから値を借りる
                Dim rNode = curNode.children(1)
                If rNode.Count > M Then
                    Me.MoveLeft(curNode, 1)
                Else
                    Me.Combine(curNode, 1, delete)
                End If
            End If
        End Sub

        ''' <summary>
        ''' 左側のノードから値を借りて、右にシフトします。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="del">削除された値のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        Private Sub MoveRight(curNode As Node, del As Integer)
            Dim left = curNode.children(del - 1)
            Dim right = curNode.children(del)

            ' 右のノードに空きを作る
            For i As Integer = right.Count - 1 To 0 Step -1
                right.values(i + 1) = right.values(i)
                right.children(i + 2) = right.children(i + 1)
            Next
            right.children(1) = right.children(0)
            right.Count += 1

            ' 親のノードの値を右側のノードに移動
            right.values(0) = curNode.values(del - 1)

            ' 左側のノードの値を親のノードに移動
            curNode.values(del - 1) = left.values(left.Count - 1)
            right.children(0) = left.children(left.Count)
            left.Count -= 1
        End Sub

        ''' <summary>
        ''' 右側のノードから値を借りて、左にシフトします。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="del">削除された値のインデックス。</param>
        ''' <remarks>このメソッドは、B木のノードを再構築するために使用されます。</remarks>
        Private Sub MoveLeft(curNode As Node, del As Integer)
            Dim left = curNode.children(del - 1)
            Dim right = curNode.children(del)

            ' 親ノードの値を左側のノードに移動
            left.Count += 1
            left.values(left.Count - 1) = curNode.values(del - 1)
            left.children(left.Count) = right.children(0)

            ' 右側のノードの値を親のノードに移動
            curNode.values(del - 1) = right.values(0)
            right.children(0) = right.children(1)
            right.Count -= 1

            ' 移動した値分、右側のノードの値と子ノードをシフト
            Dim j As Integer
            For i As Integer = 1 To right.Count
                right.values(j) = right.values(i)
                right.children(j + 1) = right.children(i + 1)
                j += 1
            Next
        End Sub

        ''' <summary>
        ''' ノードを結合するメソッド。
        ''' このメソッドは、指定されたノードを結合し、必要に応じて親ノードの値を更新します。
        ''' </summary>
        ''' <param name="curNode">現在のノード。</param>
        ''' <param name="del">削除された値のインデックス。</param>
        ''' <param name="delete">削除情報。</param>
        ''' <remarks>このメソッドは、B木のノードを再構築するために使用されます。</remarks>
        Private Sub Combine(curNode As Node, del As Integer, delete As DeleteInfo)
            Dim left = curNode.children(del - 1)
            Dim right = curNode.children(del)

            ' 左側のノードに右側のノードの値を追加
            left.Count += 1
            left.values(left.Count - 1) = curNode.values(del - 1)
            left.children(left.Count) = right.children(0)
            Dim j As Integer = left.Count
            For i As Integer = 0 To right.Count - 1
                left.values(j) = right.values(i)
                left.children(j + 1) = right.children(i + 1)
                j += 1
            Next
            left.Count += right.Count

            ' 親ノードから値を削除
            Me.RemoveItem(curNode, del - 1, delete)
        End Sub

        ''' <summary>
        ''' B木の要素を列挙するためのメソッド。
        ''' このメソッドは、B木の要素を列挙するための IEnumerator を返します。
        ''' </summary>
        ''' <returns>B木の要素を列挙する IEnumerator。</returns>
        Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
            Return New BtreeEnumerator(Me)
        End Function

        ''' <summary>
        ''' B木の要素を列挙するための非ジェネリックな IEnumerator を返すメソッド。
        ''' </summary>
        ''' <returns>B木の要素を列挙する非ジェネリックな IEnumerator。</returns>
        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        ''' <summary>
        ''' BtreeEnumeratorクラスは、B木の要素を列挙するためのIEnumerator実装です。
        ''' このクラスは、B木のノードをスタックに積み、順序通りに要素を返します。
        ''' </summary>
        ''' <typeparam name="T">対象の型。</typeparam>
        ' Btree用のIEnumerator実装
        Private NotInheritable Class BtreeEnumerator
            Implements IEnumerator(Of T)

            ' Btreeの参照
            Private ReadOnly _btree As Btree(Of T)

            ' 探索の現在の深さ
            Private _depth As Integer

            ' インデックスを保持する配列
            Private ReadOnly _index() As Integer

            ' 階層を保持する配列
            Private ReadOnly _hierarchy() As Node

            ' カレントのデータ
            Private _current As T

            ''' <summary>
            ''' このコンストラクタは、指定されたB木を使用して列挙子を初期化します。
            ''' 探索の深さと階層を計算し、最初の状態にリセットします。
            ''' </summary>
            ''' <param name="btree">対象のB木。</param>
            Public Sub New(btree As Btree(Of T))
                Me._btree = btree
                Me._depth = CInt(If(Me._btree.Count > 0, Math.Ceiling(Math.Log(Me._btree.Count) / Math.Log(btree.M * 2)), 0.0))
                Me._index = New Integer(Me._depth + 1) {}
                Me._hierarchy = New Node(Me._depth + 1) {}
                Reset()
            End Sub

            ''' <summary>
            ''' 列挙子の現在の要素を取得します。
            ''' </summary>
            ''' <returns>現在の要素。</returns>
            Public ReadOnly Property Current As T Implements IEnumerator(Of T).Current
                Get
                    Return _current
                End Get
            End Property

            ''' <summary>
            ''' 列挙子の現在の要素を取得します（非ジェネリック）。
            ''' </summary>
            ''' <returns>現在の要素。</returns>
            Private ReadOnly Property IEnumerator_Current As Object Implements IEnumerator.Current
                Get
                    Return _current
                End Get
            End Property

            ''' <summary>
            ''' 列挙子を次の要素に進めます。
            ''' このメソッドは、B木の要素を順に列挙するために使用されます。
            ''' </summary>
            ''' <returns>次の要素が存在する場合はTrue、そうでない場合はFalse。</returns>
            ''' <remarks>この実装では、スタックを使用してB木のノードを探索します。</remarks>
            Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
                SyncLock _btree._lock
                    If Me._btree._count <= 0 Then
                        Return False
                    End If

                    While True
                        If Me._hierarchy(Me._depth) IsNot Nothing Then
                            ' 現在のノードを取得
                            Dim node = Me._hierarchy(Me._depth)

                            ' 現在のインデックスを取得
                            Dim cindex = Me._index(Me._depth) And 1
                            Dim nindex = Me._index(Me._depth) >> 1

                            If cindex = 0 Then
                                ' 子ノードに移動
                                Me._depth += 1
                                Me._hierarchy(Me._depth) = node.children(nindex)
                                Me._index(Me._depth) = 0
                            ElseIf nindex < node.Count Then
                                ' 現在のノードの値を取得
                                Me._current = node.values(nindex)
                                Me._index(Me._depth) += 1
                                Return True
                            ElseIf Me._depth > 0 Then
                                ' 現在のノードの子ノードがない場合、親ノードに戻る
                                Me._depth -= 1
                                Me._index(Me._depth) += 1
                            Else
                                Exit While
                            End If
                        ElseIf Me._depth > 0 Then
                            Me._depth -= 1
                            Me._index(Me._depth) += 1
                        Else
                            Exit While
                        End If
                    End While

                    ' ここに到達した場合、列挙子は終了状態
                    Me._hierarchy(Me._depth) = Nothing
                    Me._current = CType(Nothing, T)
                    Return False
                End SyncLock
            End Function

            ''' <summary>
            ''' 列挙子を最初の位置にリセットします。
            ''' このメソッドは、列挙子を初期状態に戻し、最初の要素を指すようにします。
            ''' </summary>
            ''' <remarks>この実装では、スタックをクリアし、ルートノードを再度スタックに積みます。</remarks>
            Public Sub Reset() Implements IEnumerator.Reset
                SyncLock _btree._lock
                    Me._depth = 0
                    Me._index(Me._depth) = 0
                    Me._hierarchy(Me._depth) = Me._btree._root
                End SyncLock
            End Sub

            ''' <summary>
            ''' 列挙子を破棄します。
            ''' このメソッドは、列挙子が使用しているリソースを解放します。
            ''' </summary>
            ''' <remarks>この実装では、特にリソースを解放する必要はありません。</remarks>
            Public Sub Dispose() Implements IDisposable.Dispose
                ' 何もしない
            End Sub

        End Class

        ''' <summary>
        ''' B木のノードを表す内部クラス。
        ''' 
        ''' このクラスは、B木の各ノードを表し、キーと子ノードを格納します。
        ''' ノードは、最大 M 個の子ノードと M-1 個のキーを持つことができます。
        ''' 
        ''' このクラスは、B木の挿入や検索操作に使用されます。
        ''' </summary>
        Private NotInheritable Class Node

            ' 格納データ数
            Public Property Count As Integer

            ' ノードのキーを格納するリスト
            Public values() As T

            ' 子ノードを格納するリスト
            Public children() As Node

            ''' <summary>
            ''' コンストラクタ。
            ''' このコンストラクタは、指定されたブロックサイズに基づいてノードを初期化します。
            ''' </summary>
            ''' <param name="M">ブロックサイズ。</param>
            Public Sub New(M As Integer)
                Me.values = New T(M * 2 - 1) {}
                Me.children = New Node(M * 2) {}
            End Sub

        End Class

        ''' <summary>
        ''' B木の更新情報を表す内部クラス。
        ''' このクラスは、B木の挿入操作の結果として生成される更新情報を格納します。
        ''' </summary>
        ''' <typeparam name="T">対象の型。</typeparam>
        Private NotInheritable Class UpdateInfo

            ' 更新する値。
            Public Property Value As T

            ' 更新が行われたかどうか。
            Public Property Done As Boolean

            ' 更新後の新しく追加となったノード。
            Public Property NewNode As Node

        End Class

        ''' <summary>
        ''' B木の削除情報を表す内部クラス。
        ''' このクラスは、B木の削除操作の結果として生成される削除情報を格納します。
        ''' </summary>
        ''' <typeparam name="T">対象の型。</typeparam>
        Private NotInheritable Class DeleteInfo

            ' 削除する値。
            Public Property Value As T

            ' 削除が行われたかどうか。
            Public Property Done As Boolean

            ' 削除によりノードの再構築が必要かどうか。
            Public Property NeedRebuild As Boolean

        End Class

    End Class

End Namespace
