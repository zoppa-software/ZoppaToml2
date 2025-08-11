Imports System
Imports Xunit
Imports ZoppaTomlLibrary.Strings
Imports ZoppaTomlLibrary.Toml

Public Class TomlBlockParserTest

    <Fact>
    Sub CommentTest()
        Dim src1 = U8String.NewString("# この行は全てコメントです。
key = ""value""  # 行末までコメントです。
another = ""# これはコメントではありません""")

        ' 解析を実行
        Dim res1 = TomlParser.Parse(src1)
        Assert.Equal(2, res1.Contents.Length)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(0).Type)
        Assert.Equal(U8String.NewString("key"), res1.Contents(0).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(0).Contents(1).Str)

        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(1).Type)
        Assert.Equal(U8String.NewString("another"), res1.Contents(1).Contents(0).Str)
        Assert.Equal(U8String.NewString("""# これはコメントではありません"""), res1.Contents(1).Contents(1).Str)
    End Sub

    <Fact>
    Sub KeyValuePairTest()
        Dim src1 = U8String.NewString("key = # 無効")

        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.Parse(src1)
            End Sub
        )

        Dim src2 = U8String.NewString("first = ""Tom"" last = ""Preston-Werner"" # 無効")

        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.Parse(src2)
            End Sub
        )
    End Sub

    <Fact>
    Sub Key1Test()
        Dim src1 = U8String.NewString("key = ""value""
bare_key = ""value""
bare-key = ""value""
1234 = ""value""")

        ' 解析を実行
        Dim res1 = TomlParser.Parse(src1)
        Assert.Equal(4, res1.Contents.Length)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(0).Type)
        Assert.Equal(U8String.NewString("key"), res1.Contents(0).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(0).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(1).Type)
        Assert.Equal(U8String.NewString("bare_key"), res1.Contents(1).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(1).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(2).Type)
        Assert.Equal(U8String.NewString("bare-key"), res1.Contents(2).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(2).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(3).Type)
        Assert.Equal(U8String.NewString("1234"), res1.Contents(3).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(3).Contents(1).Str)

        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""value"""), doc1.GetExpression("key").Str)
    End Sub

    <Fact>
    Sub Key2Test()
        Dim src1 = U8String.NewString("""127.0.0.1"" = ""value""
""character encoding"" = ""value""
""ʎǝʞ"" = ""value""
'key2' = ""value""
'quoted ""value""' = ""value""")

        ' 解析を実行
        Dim res1 = TomlParser.Parse(src1)
        Assert.Equal(5, res1.Contents.Length)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(0).Type)
        Assert.Equal(U8String.NewString("""127.0.0.1"""), res1.Contents(0).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(0).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(1).Type)
        Assert.Equal(U8String.NewString("""character encoding"""), res1.Contents(1).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(1).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(2).Type)
        Assert.Equal(U8String.NewString("""ʎǝʞ"""), res1.Contents(2).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(2).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(3).Type)
        Assert.Equal(U8String.NewString("'key2'"), res1.Contents(3).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(3).Contents(1).Str)
        Assert.Equal(TomlExpressionType.Keyval, res1.Contents(4).Type)
        Assert.Equal(U8String.NewString("'quoted ""value""'"), res1.Contents(4).Contents(0).Str)
        Assert.Equal(U8String.NewString("""value"""), res1.Contents(4).Contents(1).Str)
    End Sub

    <Fact>
    Sub Key3Test()
        Dim src1 = U8String.NewString("= ""no key name""  # 不正です")
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.Parse(src1)
            End Sub
        )

        Dim src2 = U8String.NewString(""""" = ""blank""     # 可能ですがお奨めしません")
        Dim res2 = TomlParser.Parse(src2)
        Assert.Equal(1, res2.Contents.Length)
        Assert.Equal(TomlExpressionType.Keyval, res2.Contents(0).Type)
        Assert.Equal(U8String.NewString(""""""), res2.Contents(0).Contents(0).Str)
        Assert.Equal(U8String.NewString("""blank"""), res2.Contents(0).Contents(1).Str)

        Dim src3 = U8String.NewString("'' = 'blank'     # 可能ですがお奨めしません")
        Dim res3 = TomlParser.Parse(src3)
        Assert.Equal(1, res3.Contents.Length)
        Assert.Equal(TomlExpressionType.Keyval, res3.Contents(0).Type)
        Assert.Equal(U8String.NewString("''"), res3.Contents(0).Contents(0).Str)
        Assert.Equal(U8String.NewString("'blank'"), res3.Contents(0).Contents(1).Str)
    End Sub

    <Fact>
    Sub Key4Test()
        Dim src1 = U8String.NewString("name = ""Orange""
physical.color = ""orange""
physical.shape = ""round""
site.""google.com"" = true")

        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""Orange"""), doc1.GetExpression("name").Str)
        Assert.Equal(U8String.NewString("""orange"""), doc1.GetExpression("physical.color").Str)
        Assert.Equal(U8String.NewString("""round"""), doc1.GetExpression("physical.shape").Str)
        Assert.Equal(U8String.NewString("true"), doc1.GetExpression("site.""google.com""").Str)
    End Sub

    <Fact>
    Sub Key5Test()
        Dim src1 = U8String.NewString("
fruit.name = ""banana""       # これがベストプラクティスです
fruit. color = ""yellow""     # fruit.color と同じです
fruit . flavor = ""banana""   # fruit.flavor と同じです")

        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""banana"""), doc1.GetExpression("fruit.name").Str)
        Assert.Equal(U8String.NewString("""yellow"""), doc1.GetExpression("fruit.color").Str)
        Assert.Equal(U8String.NewString("""banana"""), doc1.GetExpression("fruit.flavor").Str)
    End Sub

    <Fact>
    Sub Key6Test()
        Dim src1 = U8String.NewString("
# 以下のような記述をしないでください
name = ""Tom""
name = ""Pradyun""")

        ' ドキュメント解析を実行
        Assert.Throws(Of TomlKeyDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub Key7Test()
        Dim src1 = U8String.NewString("# 以下の記述は動作しません
spelling = ""favorite""
""spelling"" = ""favourite""")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlKeyDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub Key8Test()
        Dim src1 = U8String.NewString("# これにより、キー「フルーツ」がテーブルになります。
fruit.apple.smooth = true

# したがって、以下のようにテーブルに「フルーツ」を追加できます。
fruit.orange = 2")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("true"), doc1.GetExpression("fruit.apple.smooth").Str)
        Assert.Equal(U8String.NewString("2"), doc1.GetExpression("fruit.orange").Str)

        ' テーブルの存在を確認
        Assert.True(doc1.ContainsKey("fruit"))
        Dim fruitTable = doc1.GetNode("fruit")
        Assert.NotNull(fruitTable)
        Assert.True(fruitTable.ContainsKey("apple"))
        Assert.True(fruitTable.ContainsKey("orange"))
        Assert.Equal(U8String.NewString("true"), fruitTable.GetExpression("apple.smooth").Str)
        Assert.Equal(U8String.NewString("2"), fruitTable.GetExpression("orange").Str)
    End Sub

    <Fact>
    Sub Key9Test()
        Dim src1 = U8String.NewString("# 有効ですが、推奨されません

apple.type = ""fruit""
orange.type = ""fruit""

apple.skin = ""thin""
orange.skin = ""thick""

apple.color = ""red""
orange.color = ""orange""")

        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""fruit"""), doc1.GetExpression("apple.type").Str)
        Assert.Equal(U8String.NewString("""fruit"""), doc1.GetExpression("orange.type").Str)
        Assert.Equal(U8String.NewString("""thin"""), doc1.GetExpression("apple.skin").Str)
        Assert.Equal(U8String.NewString("""thick"""), doc1.GetExpression("orange.skin").Str)
        Assert.Equal(U8String.NewString("""red"""), doc1.GetExpression("apple.color").Str)
        Assert.Equal(U8String.NewString("""orange"""), doc1.GetExpression("orange.color").Str)

        Dim src2 = U8String.NewString("3.14159 = ""pi""")
        Dim doc2 = TomlDocument.Read(src2)
        Assert.True(doc2.ContainsKey("3"))
        Assert.Equal(U8String.NewString("""pi"""), doc2.GetExpression("3.14159").Str)
    End Sub

    <Fact>
    Sub str1Test()
        Dim src1 = U8String.NewString("str = ""I'm a string. \""You can quote me\"". Name\tJos\u00E9\nLocation\tSF.""")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""I'm a string. \""You can quote me\"". Name\tJos\u00E9\nLocation\tSF."""), doc1.GetExpression("str").Str)
    End Sub

    <Fact>
    Sub str2Test()
        Dim src1 = U8String.NewString("# 以下の 3 つの文字列は、バイト単位で等価です。
str1 = ""The quick brown fox jumps over the lazy dog.""

str2 = """"""
The quick brown \


  fox jumps over \
    the lazy dog.""""""

str3 = """"""\
       The quick brown \
       fox jumps over \
       the lazy dog.\
       """"""")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal("The quick brown fox jumps over the lazy dog.", doc1.GetExpression("str1").ValueTo(Of String))
        Assert.Equal("The quick brown fox jumps over the lazy dog.", doc1.GetExpression("str2").ValueTo(Of String))
        Assert.Equal("The quick brown fox jumps over the lazy dog.", doc1.GetExpression("str3").ValueTo(Of String))
    End Sub

    <Fact>
    Sub str3Test()
        Dim src1 = U8String.NewString("str4 = """"""引用符2つ: """"。簡単ですね。""""""
# str5 = """"""引用符3つ: """"""。""""""  # 無効
str5 = """"""引用符3つ: """"\""。""""""
str6 = """"""引用符15個: """"\""""""\""""""\""""""\""""""\""。""""""

# ""This,"" she said, ""is just a pointless statement.""
str7 = """"""""This,"" she said, ""is just a pointless statement.""""""""")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal("引用符2つ: ""。簡単ですね。", doc1.GetExpression("str4").ValueTo(Of String))
        Assert.Equal("引用符3つ: """"。", doc1.GetExpression("str5").ValueTo(Of String))
        Assert.Equal("引用符15個: """"""""""""""""""""。", doc1.GetExpression("str6").ValueTo(Of String))
        Assert.Equal("""This,"" she said, ""is just a pointless statement.""", doc1.GetExpression("str7").ValueTo(Of String))
    End Sub

    <Fact>
    Sub str4Test()
        Dim src1 = U8String.NewString("# 表示されている文字列、そのものが得られます。
winpath  = 'C:\Users\nodejs\templates'
winpath2 = '\\ServerX\admin$\system32\'
quoted   = 'Tom ""Dubs"" Preston-Werner'
regex    = '<\i\c*\s*>'")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal("C:\Users\nodejs\templates", doc1.GetExpression("winpath").ValueTo(Of String))
        Assert.Equal("\\ServerX\admin$\system32\", doc1.GetExpression("winpath2").ValueTo(Of String))
        Assert.Equal("Tom ""Dubs"" Preston-Werner", doc1.GetExpression("quoted").ValueTo(Of String))
        Assert.Equal("<\i\c*\s*>", doc1.GetExpression("regex").ValueTo(Of String))
    End Sub

    <Fact>
    Sub str5Test()
        Dim src1 = U8String.NewString("regex2 = '''I [dw]on't need \d{2} apples'''
lines  = '''
生の文字列では、
最初の改行は取り除かれます。
   その他の空白は、
   保持されます。
'''")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal("I [dw]on't need \d{2} apples", doc1.GetExpression("regex2").ValueTo(Of String))
        Assert.Equal("生の文字列では、" & vbCrLf &
                     "最初の改行は取り除かれます。" & vbCrLf &
                     "   その他の空白は、" & vbCrLf &
                     "   保持されます。" & vbCrLf, doc1.GetExpression("lines").ValueTo(Of String))
    End Sub

    <Fact>
    Sub str6Test()
        Dim src1 = U8String.NewString("quot15 = '''Here are fifteen quotation marks: """"""""""""""""""""""""""""""'''

# apos15 = '''Here are fifteen apostrophes: ''''''''''''''''''  # 無効
apos15 = ""Here are fifteen apostrophes: '''''''''''''''""

# 'That,' she said, 'is still pointless.'
str = ''''That,' she said, 'is still pointless.''''")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal("Here are fifteen quotation marks: """"""""""""""""""""""""""""""", doc1.GetExpression("quot15").ValueTo(Of String))
        Assert.Equal("Here are fifteen apostrophes: '''''''''''''''", doc1.GetExpression("apos15").ValueTo(Of String))
        Assert.Equal("'That,' she said, 'is still pointless.'", doc1.GetExpression("str").ValueTo(Of String))
    End Sub

    <Fact>
    Sub Int1Test()
        Dim src1 = U8String.NewString("int1 = +99
int2 = 42
int3 = 0
int4 = -17")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(99, doc1.GetExpression("int1").ValueTo(Of Integer))
        Assert.Equal(42, doc1.GetExpression("int2").ValueTo(Of Integer))
        Assert.Equal(0, doc1.GetExpression("int3").ValueTo(Of Integer))
        Assert.Equal(-17, doc1.GetExpression("int4").ValueTo(Of Integer))
    End Sub

    <Fact>
    Sub Int2Test()
        Dim src1 = U8String.NewString("int5 = 1_000
int6 = 5_349_221
int7 = 53_49_221  # インド式命数法
int8 = 1_2_3_4_5  # 有効ですが推奨されません")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(1000, doc1.GetExpression("int5").ValueTo(Of Integer))
        Assert.Equal(5349221, doc1.GetExpression("int6").ValueTo(Of Integer))
        Assert.Equal(5349221, doc1.GetExpression("int7").ValueTo(Of Integer)) ' インド式命数法も同じ値
        Assert.Equal(12345, doc1.GetExpression("int8").ValueTo(Of Integer)) ' 12345として解釈される

        Dim src2 = U8String.NewString("# 接頭辞 `0x` が付いた 16 進数
hex1 = 0xDEADBEEF
hex2 = 0xdeadbeef
hex3 = 0xdead_beef

# 接頭辞 `0o` が付いた 8 進数
oct1 = 0o01234567
oct2 = 0o755 # Unix ファイルのパーミッションに便利

# 接頭辞 `0b` が付いた 2 進数
bin1 = 0b11010110")
        ' ドキュメント解析を実行
        Dim doc2 = TomlDocument.Read(src2)
        Assert.Equal(&HDEADBEEFUL, doc2.GetExpression("hex1").ValueTo(Of Long))
        Assert.Equal(&HDEADBEEFUL, doc2.GetExpression("hex2").ValueTo(Of Long))
        Assert.Equal(&HDEADBEEFUL, doc2.GetExpression("hex3").ValueTo(Of Long))
        Assert.Equal(342391, doc2.GetExpression("oct1").ValueTo(Of Long))
        Assert.Equal(493, doc2.GetExpression("oct2").ValueTo(Of Long)) ' 0o755は493として解釈される
        Assert.Equal(214, doc2.GetExpression("bin1").ValueTo(Of Long)) ' 0b11010110は214として解釈される
    End Sub

    <Fact>
    Sub Table1est()
        Dim src1 = U8String.NewString("[table-1]
key1 = ""some string""
key2 = 123

[table-2]
key1 = ""another string""
key2 = 456")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("table-1"))
        Assert.Equal("some string", doc1.GetExpression("table-1.key1").ValueTo(Of String))
        Assert.Equal(123, doc1.GetExpression("table-1.key2").ValueTo(Of Integer))
        Assert.True(doc1.ContainsKey("table-2"))
        Assert.Equal("another string", doc1.GetExpression("table-2.key1").ValueTo(Of String))
        Assert.Equal(456, doc1.GetExpression("table-2.key2").ValueTo(Of Integer))
    End Sub

    <Fact>
    Sub Table2Test()
        Dim src1 = U8String.NewString("[dog.""tater.man""]
type.name = ""pug""")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("dog"))
        Assert.True(doc1.GetNode("dog").ContainsKey("'tater.man'"))
        Assert.True(doc1.GetNode("dog").GetNode("'tater.man'").ContainsKey("type"))
        Assert.Equal("pug", doc1.GetExpression("dog.""tater.man"".type.name").ValueTo(Of String))
    End Sub

    <Fact>
    Sub Table3Test()
        Dim src1 = U8String.NewString("[a.b.c]            # これがベストプラクティスです
[ d.e.f ]          # [d.e.f] と同じです
[ g .  h  . i ]    # [g.h.i] と同じです
[ j . ""ʞ"" . 'l' ]  # [j.""ʞ"".'l'] と同じです")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("a.b.c"))
        Assert.True(doc1.ContainsKey("d.e.f"))
        Assert.True(doc1.ContainsKey("g.h.i"))
        Assert.True(doc1.ContainsKey("j.""ʞ"".l"))
        ' 各テーブルの存在を確認
        Dim tableABC = doc1.GetNode("a.b.c")
        Dim tableDEF = doc1.GetNode("d.e.f")
        Dim tableGHI = doc1.GetNode("g.h.i")
        Dim tableJKL = doc1.GetNode("j.""ʞ"".l")
        Assert.NotNull(tableABC)
        Assert.NotNull(tableDEF)
        Assert.NotNull(tableGHI)
        Assert.NotNull(tableJKL)
    End Sub

    <Fact>
    Sub Table4Test()
        Dim src1 = U8String.NewString("# [x] 4行目を有効にするために、
# [x.y] 1行目から3行目までを
# [x.y.z] 指定する必要はありません
[x.y.z.w]
one = 1
[x] # 後から上位テーブルを定義しても問題ありません
two = 2")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("x.y.z.w"))
        Assert.Equal(1, doc1.GetExpression("x.y.z.w.one").ValueTo(Of Integer))
        Assert.True(doc1.ContainsKey("x"))
        Assert.Equal(2, doc1.GetExpression("x.two").ValueTo(Of Integer))
        ' テーブルの存在を確認
        Dim tableXYZW = doc1.GetNode("x.y.z.w")
        Dim tableX = doc1.GetNode("x")
        Assert.NotNull(tableXYZW)
        Assert.NotNull(tableX)
    End Sub

    <Fact>
    Sub Table5Test()
        Dim src1 = U8String.NewString("# 以下のような記述をしないでください

[fruit]
apple = ""red""

[fruit]
orange = ""orange""")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub Table6Test()
        Dim src1 = U8String.NewString("# 以下のような記述もしないでください

[fruit]
apple = ""red""

[fruit.apple]
texture = ""smooth""")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlSyntaxException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub Table7Test()
        Dim src1 = U8String.NewString("# 最上位テーブルが始まります。
name = ""Fido""
breed = ""pug""

# 最上位テーブルが終了します。
[owner]
name = ""Regina Dogman""
member_since = 1999-08-04")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""Fido"""), doc1.GetExpression("name").Str)
        Assert.Equal(U8String.NewString("""pug"""), doc1.GetExpression("breed").Str)
        Assert.Equal(U8String.NewString("""Regina Dogman"""), doc1.GetExpression("owner.name").Str)
        'Assert.Equal(New Date(1999, 8, 4), doc1.GetExpression("owner.member_since").ValueTo(Of Date))

        ' テーブルの存在を確認
        Assert.True(doc1.ContainsKey("owner"))
    End Sub

    <Fact>
    Sub Table8Test()
        Dim src1 = U8String.NewString("fruit.apple.color = ""red""
# fruit という名前のテーブルを定義します
# fruit.apple という名前のテーブルを定義します

fruit.apple.taste.sweet = true
# fruit.apple.taste という名前のテーブルを定義します
# fruit と fruit.apple はすでに作成されています")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""red"""), doc1.GetExpression("fruit.apple.color").Str)
        Assert.Equal(U8String.NewString("true"), doc1.GetExpression("fruit.apple.taste.sweet").Str)
        ' テーブルの存在を確認
        Assert.True(doc1.ContainsKey("fruit"))
        Assert.True(doc1.GetNode("fruit").ContainsKey("apple"))
        Assert.True(doc1.GetNode("fruit").GetNode("apple").ContainsKey("taste"))
    End Sub

    <Fact>
    Sub Table9Test()
        Dim src1 = U8String.NewString("[fruit]
apple.color = ""red""
apple.taste.sweet = true

# [fruit.apple]  # 無効
# [fruit.apple.taste]  # 無効

[fruit.apple.texture]  # サブテーブルは追加できます
smooth = true")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""red"""), doc1.GetExpression("fruit.apple.color").Str)
        Assert.Equal(U8String.NewString("true"), doc1.GetExpression("fruit.apple.taste.sweet").Str)
        Assert.Equal(U8String.NewString("true"), doc1.GetExpression("fruit.apple.texture.smooth").Str)
        ' テーブルの存在を確認
        Assert.True(doc1.ContainsKey("fruit"))
        Assert.True(doc1.GetNode("fruit").ContainsKey("apple"))
        Assert.True(doc1.GetNode("fruit").GetNode("apple").ContainsKey("taste"))
        Assert.True(doc1.GetNode("fruit").GetNode("apple").ContainsKey("texture"))

        Dim src2 = U8String.NewString("[fruit]
apple.color = ""red""
apple.taste.sweet = true

[fruit.apple]  # 無効")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src2)
            End Sub
        )

        Dim src3 = U8String.NewString("[fruit]
apple.color = ""red""
apple.taste.sweet = true

[fruit.apple.taste]  # 無効")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src3)
            End Sub
        )
    End Sub

    <Fact>
    Sub InlineTable1Test()
        Dim src1 = U8String.NewString("name = { first = ""Tom"", last = ""Preston-Werner"" }
point = { x = 1, y = 2 }
animal = { type.name = ""pug"" }")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(U8String.NewString("""Tom"""), doc1.GetExpression("name.first").Str)
        Assert.Equal(U8String.NewString("""Preston-Werner"""), doc1.GetExpression("name.last").Str)
        Assert.Equal(U8String.NewString("1"), doc1.GetExpression("point.x").Str)
        Assert.Equal(U8String.NewString("2"), doc1.GetExpression("point.y").Str)
        Assert.Equal(U8String.NewString("""pug"""), doc1.GetExpression("animal.type.name").Str)
    End Sub

    <Fact>
    Sub InlineTable2Test()
        Dim src1 = U8String.NewString("[product]
type = { name = ""Nail"" }
type.edible = false  # 無効")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )

        Dim src2 = U8String.NewString("[product]
type.name = ""Nail""
type = { edible = false }  # 無効")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlKeyDuplicationException)(
            Sub()
                TomlDocument.Read(src2)
            End Sub
        )
    End Sub

    <Fact>
    Sub TableArray1Test()
        Dim src1 = U8String.NewString("[[products]]
name = ""Hammer""
sku = 738594937

[[products]]  # 配列内の空のテーブル

[[products]]
name = ""Nail""
sku = 284758393

color = ""gray""")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("products"))
        Dim products = doc1.GetNode("products")
        Assert.Equal(3, products.Count)
        Assert.Equal(U8String.NewString("""Hammer"""), products(0).GetExpression("name").Str)
        Assert.Equal(U8String.NewString("738594937"), products(0).GetExpression("sku").Str)
        Assert.Equal(U8String.NewString("""Nail"""), products(2).GetExpression("name").Str)
        Assert.Equal(U8String.NewString("284758393"), products(2).GetExpression("sku").Str)
        Assert.Equal(U8String.NewString("""gray"""), products(2).GetExpression("color").Str)
    End Sub

    <Fact>
    Sub TableArray2Test()
        Dim src1 = U8String.NewString("[[fruits]]
name = ""apple""

[fruits.physical]  # サブテーブル
color = ""red""
shape = ""round""

[[fruits.varieties]]  # ネストされたテーブル配列
name = ""red delicious""

[[fruits.varieties]]
name = ""granny smith""


[[fruits]]
name = ""banana""

[[fruits.varieties]]
name = ""plantain""")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("fruits"))
        Dim fruits = doc1.GetNode("fruits")
        Assert.Equal(2, fruits.Count)
        Assert.Equal(U8String.NewString("""apple"""), fruits(0).GetExpression("name").Str)
        Assert.Equal(U8String.NewString("""red"""), fruits(0).GetExpression("physical.color").Str)
        Assert.Equal(U8String.NewString("""round"""), fruits(0).GetExpression("physical.shape").Str)
        Assert.True(fruits(0).ContainsKey("varieties"))
        Dim appleVarieties = fruits(0).GetNode("varieties")
        Assert.Equal(2, appleVarieties.Count)
        Assert.Equal(U8String.NewString("""red delicious"""), appleVarieties(0).GetExpression("name").Str)
        Assert.Equal(U8String.NewString("""granny smith"""), appleVarieties(1).GetExpression("name").Str)
        Assert.Equal(U8String.NewString("""banana"""), fruits(1).GetExpression("name").Str)
        Assert.True(fruits(1).ContainsKey("varieties"))
        Dim bananaVarieties = fruits(1).GetNode("varieties")
        Assert.Equal(1, bananaVarieties.Count)
        Assert.Equal(U8String.NewString("""plantain"""), bananaVarieties(0).GetExpression("name").Str)
    End Sub

    <Fact>
    Sub TableArray3Test()
        Dim src1 = U8String.NewString("# 無効な TOML ドキュメント
[fruit.physical]  # サブテーブルですが、どの親要素に属すべきでしょうか?
color = ""red""
shape = ""round""

[[fruit]]  # パーサーは ""fruit"" がテーブルではなく配列であることを発見したときにエラーをスローする必要があります
name = ""apple""")
        Assert.Throws(Of TomlSyntaxException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub TableArray4Test()
        Dim src1 = U8String.NewString("# 無効な TOML ドキュメント
fruits = []

[[fruits]] # 許可されていません")
        Assert.Throws(Of TomlSyntaxException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub TableArray5Test()
        Dim src1 = U8String.NewString("# 無効な TOML ドキュメント
[[fruits]]
name = ""apple""

[[fruits.varieties]]
name = ""red delicious""

# 無効: このテーブルは前のテーブルの配列と競合しています
[fruits.varieties]
name = ""granny smith""")
        Assert.Throws(Of TomlTableDuplicationException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub TableArray6Test()
        Dim src1 = U8String.NewString("# 無効な TOML ドキュメント
[[fruits]]
name = ""apple""

[[fruits.varieties]]
name = ""red delicious""

[fruits.physical]
color = ""red""
shape = ""round""

# 無効: このテーブルの配列は前のテーブルと競合します
[[fruits.physical]]
color = ""green""")
        Assert.Throws(Of TomlSyntaxException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub TableArray7Test()
        Dim src1 = U8String.NewString("points = [ { x = 1, y = 2, z = 3 },
           { x = 7, y = 8, z = 9 },
           { x = 2, y = 4, z = 8 } ]")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(1, doc1.GetNode("points")(0).GetExpression("x").ValueTo(Of Integer)())
        Assert.Equal(2, doc1.GetNode("points")(0).GetExpression("y").ValueTo(Of Integer)())
        Assert.Equal(3, doc1.GetNode("points")(0).GetExpression("z").ValueTo(Of Integer)())
        Assert.Equal(7, doc1.GetNode("points")(1).GetExpression("x").ValueTo(Of Integer)())
        Assert.Equal(8, doc1.GetNode("points")(1).GetExpression("y").ValueTo(Of Integer)())
        Assert.Equal(9, doc1.GetNode("points")(1).GetExpression("z").ValueTo(Of Integer)())
        Assert.Equal(2, doc1.GetNode("points")(2).GetExpression("x").ValueTo(Of Integer)())
        Assert.Equal(4, doc1.GetNode("points")(2).GetExpression("y").ValueTo(Of Integer)())
        Assert.Equal(8, doc1.GetNode("points")(2).GetNode("z").ValueTo(Of Integer)())
    End Sub

    <Fact>
    Sub Array1Test()
        Dim src1 = U8String.NewString("integers = [ 1, 2, 3 ]
colors = [ ""red"", ""yellow"", ""green"" ]
nested_arrays_of_ints = [ [ 1, 2 ], [3, 4, 5] ]
nested_mixed_array = [ [ 1, 2 ], [""a"", ""b"", ""c""] ]
string_array = [ ""all"", 'strings', """"""are the same"""""", '''type''' ]

# 異なるデータ型の値を混在させることができます
numbers = [ 0.1, 0.2, 0.5, 1, 2, 5 ]
contributors = [
  ""Foo Bar <foo@example.com>"",
  { name = ""Baz Qux"", email = ""bazqux@example.com"", url = ""https://example.com/bazqux"" }
]")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(doc1.ContainsKey("integers"))
        Dim integers = doc1.GetNode("integers")
        Assert.Equal(3, integers.Count)
        Assert.Equal(1, integers(0).GetExpression().ValueTo(Of Integer))
        Assert.Equal(2, integers(1).GetExpression().ValueTo(Of Integer))
        Assert.Equal(3, integers(2).GetExpression().ValueTo(Of Integer))
        Assert.True(doc1.ContainsKey("colors"))
        Dim colors = doc1.GetNode("colors")
        Assert.Equal(3, colors.Count)
        Assert.Equal(U8String.NewString("""red"""), colors(0).GetExpression().Str)
        Assert.Equal(U8String.NewString("""yellow"""), colors(1).GetExpression().Str)
        Assert.Equal(U8String.NewString("""green"""), colors(2).GetExpression().Str)
        Assert.True(doc1.ContainsKey("nested_arrays_of_ints"))
        Dim nestedArraysOfInts = doc1.GetNode("nested_arrays_of_ints")
        Assert.Equal(2, nestedArraysOfInts.Count)
        Assert.Equal(1, nestedArraysOfInts(0)(0).GetExpression().ValueTo(Of Integer))
        Assert.Equal(2, nestedArraysOfInts(0)(1).GetExpression().ValueTo(Of Integer))
        Assert.Equal(3, nestedArraysOfInts(1)(0).GetExpression().ValueTo(Of Integer))
        Assert.Equal(4, nestedArraysOfInts(1)(1).GetExpression().ValueTo(Of Integer))
        Assert.Equal(5, nestedArraysOfInts(1)(2).GetExpression().ValueTo(Of Integer))
        Assert.True(doc1.ContainsKey("nested_mixed_array"))
        Dim nestedMixedArray = doc1.GetNode("nested_mixed_array")
        Assert.Equal(2, nestedMixedArray.Count)
        Assert.Equal(1, nestedMixedArray(0)(0).GetExpression().ValueTo(Of Integer))
        Assert.Equal(2, nestedMixedArray(0)(1).GetExpression().ValueTo(Of Integer))
        Assert.Equal(U8String.NewString("""a"""), nestedMixedArray(1)(0).GetExpression().Str)
        Assert.Equal(U8String.NewString("""b"""), nestedMixedArray(1)(1).GetExpression().Str)
        Assert.Equal(U8String.NewString("""c"""), nestedMixedArray(1)(2).GetExpression().Str)
        Assert.True(doc1.ContainsKey("string_array"))
        Dim stringArray = doc1.GetNode("string_array")
        Assert.Equal(4, stringArray.Count)
        Assert.Equal("all", stringArray(0).GetExpression().ValueTo(Of String))
        Assert.Equal("strings", stringArray(1).GetExpression().ValueTo(Of String))
        Assert.Equal("are the same", stringArray(2).GetExpression().ValueTo(Of String))
        Assert.Equal("type", stringArray(3).GetExpression().ValueTo(Of String))
        Assert.True(doc1.ContainsKey("numbers"))
        Dim numbers = doc1.GetNode("numbers")
        Assert.Equal(6, numbers.Count)
        Assert.Equal(0.1, numbers(0).GetExpression().ValueTo(Of Double))
        Assert.Equal(0.2, numbers(1).GetExpression().ValueTo(Of Double))
        Assert.Equal(0.5, numbers(2).GetExpression().ValueTo(Of Double))
        Assert.Equal(1, numbers(3).GetExpression().ValueTo(Of Integer))
        Assert.Equal(2, numbers(4).GetExpression().ValueTo(Of Integer))
        Assert.Equal(5, numbers(5).GetExpression().ValueTo(Of Integer))
        Assert.True(doc1.ContainsKey("contributors"))
        Dim contributors = doc1.GetNode("contributors")
        Assert.Equal(2, contributors.Count)
        Assert.Equal("Foo Bar <foo@example.com>", contributors(0).GetExpression().ValueTo(Of String))
        Assert.Equal("Baz Qux", contributors(1).GetExpression("name").ValueTo(Of String))
        Assert.Equal("bazqux@example.com", contributors(1).GetExpression("email").ValueTo(Of String))
        Assert.Equal("https://example.com/bazqux", contributors(1).GetExpression("url").ValueTo(Of String))
    End Sub

    <Fact>
    Sub Float1Test()
        Dim src1 = U8String.NewString("# 小数部
flt1 = +1.0
flt2 = 3.1415
flt3 = -0.01

# 指数部
flt4 = 5e+22
flt5 = 1e06
flt6 = -2E-2

# 少数部と指数部の両方
flt7 = 6.626e-34")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(1.0, doc1.GetExpression("flt1").ValueTo(Of Double))
        Assert.Equal(3.1415, doc1.GetExpression("flt2").ValueTo(Of Double))
        Assert.Equal(-0.01, doc1.GetExpression("flt3").ValueTo(Of Double))
        Assert.True(Math.Abs(5.0E+22 - doc1.GetExpression("flt4").ValueTo(Of Double)) < 0.000001)
        Assert.True(Math.Abs(1000000.0 - doc1.GetExpression("flt5").ValueTo(Of Double)) < 0.000001)
        Assert.True(Math.Abs(-0.02 - doc1.GetExpression("flt6").ValueTo(Of Double)) < 0.000001)
        Assert.True(Math.Abs(6.626E-34 - doc1.GetExpression("flt7").ValueTo(Of Double)) < 0.000001)
    End Sub

    <Fact>
    Sub Float2Test()
        Dim src1 = U8String.NewString("# 無効な浮動小数点数
invalid_float_1 = .7
invalid_float_2 = 7.
invalid_float_3 = 3.e+20")
        ' ドキュメント解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlDocument.Read(src1)
            End Sub
        )
    End Sub

    <Fact>
    Sub Float3Test()
        Dim src1 = U8String.NewString("flt8 = 224_617.445_991_228")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.True(Math.Abs(224617.445991228 - doc1.GetExpression("flt8").ValueTo(Of Double)) < 0.000001)
    End Sub

    <Fact>
    Sub Float4Test()
        Dim src1 = U8String.NewString("# 無限大
sf1 = inf  # 正の無限大
sf2 = +inf # 正の無限大
sf3 = -inf # 負の無限大

# 非数 (NaN)
sf4 = nan  # 実際の sNaN/qNaN エンコーディングは実装に依存します
sf5 = +nan # `nan` と等価
sf6 = -nan # 有効で、実際のエンコーディングは実装に依存します")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(Double.PositiveInfinity, doc1.GetExpression("sf1").ValueTo(Of Double))
        Assert.Equal(Double.PositiveInfinity, doc1.GetExpression("sf2").ValueTo(Of Double))
        Assert.Equal(Double.NegativeInfinity, doc1.GetExpression("sf3").ValueTo(Of Double))
        Assert.True(Double.IsNaN(doc1.GetExpression("sf4").ValueTo(Of Double)))
        Assert.True(Double.IsNaN(doc1.GetExpression("sf5").ValueTo(Of Double)))
        Assert.True(Double.IsNaN(doc1.GetExpression("sf6").ValueTo(Of Double)))
    End Sub

    <Fact>
    Sub BoolTest()
        Dim src1 = U8String.NewString("bool1 = true
bool2 = false")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(True, doc1.GetExpression("bool1").ValueTo(Of Boolean))
        Assert.Equal(False, doc1.GetExpression("bool2").ValueTo(Of Boolean))
    End Sub

    <Fact>
    Sub DateTime1Test()
        Dim src1 = U8String.NewString("lt1 = 07:32:00
lt2 = 00:32:00.999999")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(New TimeSpan(7, 32, 0), doc1.GetExpression("lt1").ValueTo(Of TimeSpan))
        Assert.Equal(New TimeSpan(0, 0, 32, 0, 999), doc1.GetExpression("lt2").ValueTo(Of TimeSpan))
    End Sub

    <Fact>
    Sub DateTime2Test()
        Dim src1 = U8String.NewString("ld1 = 1979-05-27")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(New Date(1979, 5, 27), doc1.GetExpression("ld1").ValueTo(Of Date))
    End Sub

    <Fact>
    Sub DateTime3Test()
        Dim src1 = U8String.NewString("ldt1 = 1979-05-27T07:32:00
ldt2 = 1979-05-27T00:00:32.999999")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(New Date(1979, 5, 27, 7, 32, 0), doc1.GetExpression("ldt1").ValueTo(Of Date))
        Assert.Equal(New Date(1979, 5, 27, 0, 0, 32, 999), doc1.GetExpression("ldt2").ValueTo(Of Date))
    End Sub

    <Fact>
    Sub DateTime4Test()
        Dim src1 = U8String.NewString("odt4 = 1979-05-27 07:32:00Z")
        Dim doc1 = TomlDocument.Read(src1)
        Assert.Equal(New DateTimeOffset(1979, 5, 27, 7, 32, 0, TimeSpan.Zero), doc1.GetExpression("odt4").ValueTo(Of DateTimeOffset))

        Dim src2 = U8String.NewString("odt1 = 1979-05-27T07:32:00Z
odt2 = 1979-05-27T00:32:00-07:00
odt3 = 1979-05-27T00:32:00.999999-07:00")
        Dim doc2 = TomlDocument.Read(src2)
        Assert.Equal(New DateTimeOffset(1979, 5, 27, 7, 32, 0, TimeSpan.Zero), doc2.GetExpression("odt1").ValueTo(Of DateTimeOffset))
        Assert.Equal(New DateTimeOffset(1979, 5, 27, 0, 32, 0, TimeSpan.FromHours(-7)), doc2.GetExpression("odt2").ValueTo(Of DateTimeOffset))
        Assert.Equal(New DateTimeOffset(1979, 5, 27, 0, 32, 0, 999, TimeSpan.FromHours(-7)), doc2.GetExpression("odt3").ValueTo(Of DateTimeOffset))
    End Sub

    <Fact>
    Sub InlineTableArray1Test()
        Dim src1 = U8String.NewString("inline_table_array = [ { x = 1, y = 2 }, { x = 3, y = 4 } ]")
        ' ドキュメント解析を実行
        Dim doc1 = TomlDocument.Read(src1)
        Dim ita = doc1.GetNode("inline_table_array")
        Assert.Equal(2, ita.Count)
        Assert.Equal(1, ita(0).GetExpression("x").ValueTo(Of Integer))
        Assert.Equal(2, ita(0).GetExpression("y").ValueTo(Of Integer))
        Assert.Equal(3, ita(1).GetExpression("x").ValueTo(Of Integer))
        Assert.Equal(4, ita(1).GetExpression("y").ValueTo(Of Integer))

        Dim src2 = U8String.NewString("inline_inline_table = { a = { x = 1, y = 2 }, b = { x = 3, y = 4 }, c = { x = 5, y = 6 } }")
        ' ドキュメント解析を実行
        Dim doc2 = TomlDocument.Read(src2)
        Dim iit = doc2.GetNode("inline_inline_table")
        Assert.True(iit.ContainsKey("a"))
        Assert.True(iit.ContainsKey("b"))
        Assert.True(iit.ContainsKey("c"))
        Assert.Equal(1, iit("a")("x").ValueTo(Of Integer))
        Assert.Equal(2, iit("a")("y").ValueTo(Of Integer))
        Assert.Equal(3, iit("b")("x").ValueTo(Of Integer))
        Assert.Equal(4, iit("b")("y").ValueTo(Of Integer))
        Assert.Equal(5, iit("c")("x").ValueTo(Of Integer))
        Assert.Equal(6, iit("c")("y").ValueTo(Of Integer))
    End Sub

End Class
