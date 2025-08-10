Imports System
Imports Xunit
Imports ZoppaTomlLibrary.Strings
Imports ZoppaTomlLibrary.Toml

Public Class TomlParserTest

    <Fact>
    Sub ParseUnquotedKeyTest()
        Dim source = U8String.NewString("example_key []")
        Dim iter = source.GetIterator()

        ' 解析を実行
        Dim res1 = TomlParser.ParseUnquotedKey(source, iter)
        Assert.Equal(TomlExpressionType.UnquotedKey, res1.Type)
        Assert.Equal("example_key", res1.Str.ToString())

        Dim res2 = TomlParser.ParseUnquotedKey(source, iter)
        Assert.Equal(TomlExpressionType.None, res2.Type)
    End Sub

    <Fact>
    Sub ParseUnsignedDecIntTest()
        Dim source = U8String.NewString("12345_67890")
        Dim iter = source.GetIterator()

        ' 解析を実行
        Dim res = TomlParser.ParseInteger(source, iter)
        Assert.Equal(TomlExpressionType.DecInt, res.Type)
        Assert.Equal("12345_67890", res.Str.ToString())

        ' 次の文字が数字でない場合は終了
        Assert.False(iter.HasNext())
    End Sub

    <Fact>
    Sub ParseUnsignedDecIntTest2()
        Dim source = U8String.NewString("0")
        Dim iter = source.GetIterator()

        ' 解析を実行
        Dim res = TomlParser.ParseInteger(source, iter)
        Assert.Equal(TomlExpressionType.DecInt, res.Type)
        Assert.Equal("0", res.Str.ToString())

        ' 次の文字が数字でない場合は終了
        Assert.False(iter.HasNext())
    End Sub

    <Fact>
    Sub ParseHexIntTest()
        Dim source = U8String.NewString("0x1A2B3C")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseHexInt(source, iter)
        Assert.Equal(TomlExpressionType.HexInt, res.Type)
        Assert.Equal("0x1A2B3C", res.Str.ToString())
        ' 次の文字が数字でない場合は終了
        Assert.False(iter.HasNext())
    End Sub

    <Fact>
    Sub ParseWsTest()
        Dim source = U8String.NewString($"   {vbTab}{vbCr}{vbCrLf}")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseWs(source, iter)
        Assert.Equal(TomlExpressionType.Ws, res.Type)
        Assert.Equal($"   {vbTab}", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseCommentTest()
        Dim source = U8String.NewString("# This is a comment")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseComment(source, iter)
        Assert.Equal(TomlExpressionType.Comment, res.Type)
        Assert.Equal("# This is a comment", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseTrueLiteralTest()
        Dim source = U8String.NewString("true")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseBoolean(source, iter)
        Assert.Equal(TomlExpressionType.BooleanLiteral, res.Type)
        Assert.Equal("true", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseFalseLiteralTest()
        Dim source = U8String.NewString("false")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseBoolean(source, iter)
        Assert.Equal(TomlExpressionType.BooleanLiteral, res.Type)
        Assert.Equal("false", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseEscapedTest()
        Dim src1 = U8String.NewString("\""")
        Dim res1 = TomlParser.MatchEscaped(src1.GetIterator())
        Assert.True(res1)

        Dim src2 = U8String.NewString("\\")
        Dim res2 = TomlParser.MatchEscaped(src2.GetIterator())
        Assert.True(res2)

        Dim src3 = U8String.NewString("\u1234")
        Dim res3 = TomlParser.MatchEscaped(src3.GetIterator())
        Assert.True(res3)

        Dim src4 = U8String.NewString("\U0001F600")
        Dim res4 = TomlParser.MatchEscaped(src4.GetIterator())
        Assert.True(res4)

        Dim src5 = U8String.NewString("\n")
        Dim res5 = TomlParser.MatchEscaped(src5.GetIterator())
        Assert.True(res5)

        Dim src6 = U8String.NewString("\t")
        Dim res6 = TomlParser.MatchEscaped(src6.GetIterator())
        Assert.True(res6)

        Dim src7 = U8String.NewString("\r")
        Dim res7 = TomlParser.MatchEscaped(src7.GetIterator())
        Assert.True(res7)

        Dim src8 = U8String.NewString("\f")
        Dim res8 = TomlParser.MatchEscaped(src8.GetIterator())
        Assert.True(res8)

        Dim src9 = U8String.NewString("\b")
        Dim res9 = TomlParser.MatchEscaped(src9.GetIterator())
        Assert.True(res9)

        Dim src10 = U8String.NewString("\v")
        Dim res10 = TomlParser.MatchEscaped(src10.GetIterator())
        Assert.False(res10)
    End Sub

    <Fact>
    Sub ParseBasicCharTest()
        Dim source = U8String.NewString("12\nA")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.MatchBasicChar(iter)
        Assert.True(res) ' 1
        res = TomlParser.MatchBasicChar(iter)
        Assert.True(res) ' 2
        res = TomlParser.MatchEscaped(iter)
        Assert.True(res) ' \n
        res = TomlParser.MatchBasicChar(iter)
        Assert.True(res) ' A
    End Sub

    <Fact>
    Sub ParseBasicStringTest()
        Dim source = U8String.NewString("""Hello\tWorld"" ")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseBasicString(source, iter)
        Assert.Equal(TomlExpressionType.BasicString, res.Type)
        Assert.Equal("""Hello\tWorld""", res.Str.ToString())

        Dim errSource = U8String.NewString("""Hello\tWorld ")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseBasicString(errSource, errIter)
            End Sub
        )
    End Sub

    <Fact>
    Sub ParseLiteralStringTest()
        Dim source = U8String.NewString("'Hello World'")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseLiteralString(source, iter)
        Assert.Equal(TomlExpressionType.LiteralString, res.Type)
        Assert.Equal("'Hello World'", res.Str.ToString())

        Dim errSource = U8String.NewString("'Hello World")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseLiteralString(errSource, errIter)
            End Sub
        )
    End Sub

    <Fact>
    Sub ParseQuotedKeyTest()
        Dim source = U8String.NewString("""example_key""'example_key'example_key")
        Dim iter = source.GetIterator()

        ' 解析を実行
        Dim res1 = TomlParser.ParseQuotedKey(source, iter)
        Assert.Equal(TomlExpressionType.QuotedKey, res1.Type)
        Assert.Equal("""example_key""", res1.Str.ToString())

        Dim res2 = TomlParser.ParseQuotedKey(source, iter)
        Assert.Equal(TomlExpressionType.QuotedKey, res2.Type)
        Assert.Equal("'example_key'", res2.Str.ToString())

        Dim res3 = TomlParser.ParseQuotedKey(source, iter)
        Assert.Equal(TomlExpressionType.None, res3.Type)
    End Sub

    <Fact>
    Sub ParseArrayTableTest()
        Dim source = U8String.NewString("[[example_table]]")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseArrayTable(source, iter)
        Assert.Equal(TomlExpressionType.ArrayTable, res.Type)
        Assert.Equal("[[example_table]]", res.Str.ToString())

        Dim errSource = U8String.NewString("[[example_table")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseArrayTable(errSource, errIter)
            End Sub
        )

        Dim src1 = U8String.NewString("[[fruits.varieties]]")
        Dim iter1 = src1.GetIterator()
        Dim res1 = TomlParser.ParseArrayTable(src1, iter1)
        Assert.Equal(TomlExpressionType.ArrayTable, res1.Type)
        Assert.Equal("[[fruits.varieties]]", res1.Str.ToString())
    End Sub

    <Fact>
    Sub ParseInlineTableTest()
        Dim source = U8String.NewString("{example_key = 123, another_key = 456}")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseInlineTable(source, iter)
        Assert.Equal(TomlExpressionType.InlineTable, res.Type)
        Assert.Equal("{example_key = 123, another_key = 456}", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseIntegerTest()
        Dim source = U8String.NewString("12345")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseInteger(source, iter)
        Assert.Equal(TomlExpressionType.DecInt, res.Type)
        Assert.Equal("12345", res.Str.ToString())

        Dim src1 = U8String.NewString("0x1A2B3C")
        Dim iter1 = src1.GetIterator()
        Dim res1 = TomlParser.ParseHexInt(src1, iter1)
        Assert.Equal(TomlExpressionType.HexInt, res1.Type)
        Assert.Equal("0x1A2B3C", res1.Str.ToString())

        Dim src2 = U8String.NewString("0o755")
        Dim iter2 = src2.GetIterator()
        Dim res2 = TomlParser.ParseOctInt(src2, iter2)
        Assert.Equal(TomlExpressionType.OctInt, res2.Type)
        Assert.Equal("0o755", res2.Str.ToString())

        Dim src3 = U8String.NewString("0b101010")
        Dim iter3 = src3.GetIterator()
        Dim res3 = TomlParser.ParseBinInt(src3, iter3)
        Assert.Equal(TomlExpressionType.BinInt, res3.Type)
        Assert.Equal("0b101010", res3.Str.ToString())

        Dim src4 = U8String.NewString("+99")
        Dim iter4 = src4.GetIterator()
        Dim res4 = TomlParser.ParseInteger(src4, iter4)
        Assert.Equal(TomlExpressionType.DecInt, res4.Type)
        Assert.Equal("+99", res4.Str.ToString())

        Dim src5 = U8String.NewString("-42")
        Dim iter5 = src5.GetIterator()
        Dim res5 = TomlParser.ParseInteger(src5, iter5)
        Assert.Equal(TomlExpressionType.DecInt, res5.Type)

        Dim src6 = U8String.NewString("0")
        Dim iter6 = src6.GetIterator()
        Dim res6 = TomlParser.ParseInteger(src6, iter6)
        Assert.Equal(TomlExpressionType.DecInt, res6.Type)
        Assert.Equal("0", res6.Str.ToString())

        Dim src7 = U8String.NewString("123_456")
        Dim iter7 = src7.GetIterator()
        Dim res7 = TomlParser.ParseInteger(src7, iter7)
        Assert.Equal(TomlExpressionType.DecInt, res7.Type)

        Dim src8 = U8String.NewString("-0")
        Dim iter8 = src8.GetIterator()
        Dim res8 = TomlParser.ParseInteger(src8, iter8)
        Assert.Equal(TomlExpressionType.DecInt, res8.Type)
        Assert.Equal("-0", res8.Str.ToString())
    End Sub

    <Fact>
    Sub ParseSpecialFloatTest()
        Dim source = U8String.NewString("inf")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseSpecialFloat(source, iter)
        Assert.Equal(TomlExpressionType.Inf, res.Type)
        Assert.Equal("inf", res.Str.ToString())

        Dim source2 = U8String.NewString("nan")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseSpecialFloat(source2, iter2)
        Assert.Equal(TomlExpressionType.Nan, res2.Type)
        Assert.Equal("nan", res2.Str.ToString())

        Dim source3 = U8String.NewString("-inf")
        Dim iter3 = source3.GetIterator()
        ' 解析を実行
        Dim res3 = TomlParser.ParseSpecialFloat(source3, iter3)
        Assert.Equal(TomlExpressionType.Inf, res3.Type)
        Assert.Equal("-inf", res3.Str.ToString())

        Dim source4 = U8String.NewString("+inf")
        Dim iter4 = source4.GetIterator()
        ' 解析を実行
        Dim res4 = TomlParser.ParseSpecialFloat(source4, iter4)
        Assert.Equal(TomlExpressionType.Inf, res4.Type)

        Dim source5 = U8String.NewString("+nan")
        Dim iter5 = source5.GetIterator()
        ' 解析を実行
        Dim res5 = TomlParser.ParseSpecialFloat(source5, iter5)
        Assert.Equal(TomlExpressionType.Nan, res5.Type)
        Assert.Equal("+nan", res5.Str.ToString())

        Dim source6 = U8String.NewString("-nan")
        Dim iter6 = source6.GetIterator()
        ' 解析を実行
        Dim res6 = TomlParser.ParseSpecialFloat(source6, iter6)
        Assert.Equal(TomlExpressionType.Nan, res6.Type)
        Assert.Equal("-nan", res6.Str.ToString())
    End Sub

    <Fact>
    Sub ParseFloatTest()
        Dim source = U8String.NewString("3.14")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseFloat(source, iter)
        Assert.Equal(TomlExpressionType.Float, res.Type)
        Assert.Equal("3.14", res.Str.ToString())

        Dim source2 = U8String.NewString("2.71828e10")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseFloat(source2, iter2)
        Assert.Equal(TomlExpressionType.Float, res2.Type)
        Assert.Equal("2.71828e10", res2.Str.ToString())

        Dim source3 = U8String.NewString("+1.0")
        Dim iter3 = source3.GetIterator()
        ' 解析を実行
        Dim res3 = TomlParser.ParseFloat(source3, iter3)
        Assert.Equal(TomlExpressionType.Float, res3.Type)
        Assert.Equal("+1.0", res3.Str.ToString())

        Dim source4 = U8String.NewString("3.1415")
        Dim iter4 = source4.GetIterator()
        ' 解析を実行
        Dim res4 = TomlParser.ParseFloat(source4, iter4)
        Assert.Equal(TomlExpressionType.Float, res4.Type)
        Assert.Equal("3.1415", res4.Str.ToString())

        Dim source5 = U8String.NewString("-2.5")
        Dim iter5 = source5.GetIterator()
        ' 解析を実行
        Dim res5 = TomlParser.ParseFloat(source5, iter5)
        Assert.Equal(TomlExpressionType.Float, res5.Type)
        Assert.Equal("-2.5", res5.Str.ToString())

        Dim source6 = U8String.NewString("5e+22")
        Dim iter6 = source6.GetIterator()
        ' 解析を実行
        Dim res6 = TomlParser.ParseFloat(source6, iter6)
        Assert.Equal(TomlExpressionType.Float, res6.Type)
        Assert.Equal("5e+22", res6.Str.ToString())

        Dim source7 = U8String.NewString("1.0e-3")
        Dim iter7 = source7.GetIterator()
        ' 解析を実行
        Dim res7 = TomlParser.ParseFloat(source7, iter7)
        Assert.Equal(TomlExpressionType.Float, res7.Type)
        Assert.Equal("1.0e-3", res7.Str.ToString())

        Dim source8 = U8String.NewString("-2E-2")
        Dim iter8 = source8.GetIterator()
        ' 解析を実行
        Dim res8 = TomlParser.ParseFloat(source8, iter8)
        Assert.Equal(TomlExpressionType.Float, res8.Type)
        Assert.Equal("-2E-2", res8.Str.ToString())
    End Sub

    <Fact>
    Sub ParseParseTable()
        Dim source = U8String.NewString("[example_table]")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseTable(source, iter)
        Assert.Equal(TomlExpressionType.Table, res.Type)
        Assert.Equal("[example_table]", res.Str.ToString())

        Dim errSource = U8String.NewString("[example_table")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseTable(errSource, errIter)
            End Sub
        )

        Dim src1 = U8String.NewString("[dog.""tater.man""]")
        Dim iter1 = src1.GetIterator()
        Dim res1 = TomlParser.ParseTable(src1, iter1)
        Assert.Equal(TomlExpressionType.Table, res1.Type)
        Assert.Equal("[dog.""tater.man""]", res1.Str.ToString())

        Dim src2 = U8String.NewString("[[fruits.varieties]]")
        Dim iter2 = src2.GetIterator()
        Dim res2 = TomlParser.ParseTable(src2, iter2)
        Assert.Equal(TomlExpressionType.ArrayTable, res2.Type)
        Assert.Equal("[[fruits.varieties]]", res2.Str.ToString())
    End Sub

    <Fact>
    Sub ParseArrayValuesTest()
        Dim source = U8String.NewString("
# 値1
1, 
# 値2
2,
# 値3
3
")
        Dim iter = source.GetIterator()

        ' 解析を実行
        Dim expressions As New List(Of TomlExpression)()
        Dim res = TomlParser.ParseArrayValues(source, iter, expressions)
        Assert.Equal(TomlExpressionType.ArrayValues, res.Type)
        Assert.Equal("
# 値1
1, 
# 値2
2,
# 値3
3
", res.Str.ToString())
        Assert.Equal(3, expressions.Count)
        Assert.Equal(TomlExpressionType.DecInt, expressions(0).Type)
        Assert.Equal("1", expressions(0).Str.ToString())
        Assert.Equal(TomlExpressionType.DecInt, expressions(1).Type)
        Assert.Equal("2", expressions(1).Str.ToString())
        Assert.Equal(TomlExpressionType.DecInt, expressions(2).Type)
        Assert.Equal("3", expressions(2).Str.ToString())
    End Sub

    <Fact>
    Sub ParseArrayTest()
        Dim source = U8String.NewString("[1, 2, 3]")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseArray(source, iter)
        Assert.Equal(TomlExpressionType.Array, res.Type)
        Assert.Equal("[1, 2, 3]", res.Str.ToString())
        Dim errSource = U8String.NewString("[1, 2, 3")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseArray(errSource, errIter)
            End Sub
        )
    End Sub

    <Fact>
    Sub ParseLocalDateTest()
        Dim source = U8String.NewString("2023-10-01")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseLocalDate(source, iter)
        Assert.Equal(TomlExpressionType.LocalDate, res.Type)
        Assert.Equal("2023-10-01", res.Str.ToString())
    End Sub

    <Fact>
    Sub ParseLocalTimeTest()
        Dim source = U8String.NewString("12:34:56")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseLocalTime(source, iter)
        Assert.Equal(TomlExpressionType.LocalTime, res.Type)
        Assert.Equal("12:34:56", res.Str.ToString())

        Dim source2 = U8String.NewString("12:34:56.789")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseLocalTime(source2, iter2)
        Assert.Equal(TomlExpressionType.LocalTime, res2.Type)
        Assert.Equal("12:34:56.789", res2.Str.ToString())
    End Sub

    <Fact>
    Sub ParseLocalDateTimeTest()
        Dim source = U8String.NewString("2023-10-01T12:34:56")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseLocalDateTime(source, iter)
        Assert.Equal(TomlExpressionType.LocalDateTime, res.Type)
        Assert.Equal("2023-10-01T12:34:56", res.Str.ToString())

        Dim source2 = U8String.NewString("2023-10-01T12:34:56.789")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseLocalDateTime(source2, iter2)
        Assert.Equal(TomlExpressionType.LocalDateTime, res2.Type)
        Assert.Equal("2023-10-01T12:34:56.789", res2.Str.ToString())
    End Sub

    <Fact>
    Sub ParseOffsetDateTimeTest()
        Dim source = U8String.NewString("2023-10-01T12:34:56Z")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseOffsetDateTime(source, iter)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res.Type)
        Assert.Equal("2023-10-01T12:34:56Z", res.Str.ToString())

        Dim source2 = U8String.NewString("2023-10-01T12:34:56+09:00")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseOffsetDateTime(source2, iter2)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res2.Type)
        Assert.Equal("2023-10-01T12:34:56+09:00", res2.Str.ToString())
    End Sub

    <Fact>
    Sub ParseDateTimeTest()
        Dim source = U8String.NewString("2023-10-01T12:34:56Z")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseDateTime(source, iter)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res.Type)
        Assert.Equal("2023-10-01T12:34:56Z", res.Str.ToString())

        Dim source2 = U8String.NewString("2023-10-01T12:34:56+09:00")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseDateTime(source2, iter2)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res2.Type)
        Assert.Equal("2023-10-01T12:34:56+09:00", res2.Str.ToString())

        Dim source3 = U8String.NewString("2023-10-01T12:34:56.789Z")
        Dim iter3 = source3.GetIterator()
        ' 解析を実行
        Dim res3 = TomlParser.ParseDateTime(source3, iter3)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res3.Type)
        Assert.Equal("2023-10-01T12:34:56.789Z", res3.Str.ToString())

        Dim source4 = U8String.NewString("2023-10-01T12:34:56.789+09:00")
        Dim iter4 = source4.GetIterator()
        ' 解析を実行
        Dim res4 = TomlParser.ParseDateTime(source4, iter4)
        Assert.Equal(TomlExpressionType.OffsetDateTime, res4.Type)
        Assert.Equal("2023-10-01T12:34:56.789+09:00", res4.Str.ToString())

        Dim source5 = U8String.NewString("2023-10-01T12:34:56.789")
        Dim iter5 = source5.GetIterator()
        ' 解析を実行
        Dim res5 = TomlParser.ParseDateTime(source5, iter5)
        Assert.Equal(TomlExpressionType.LocalDateTime, res5.Type)
        Assert.Equal("2023-10-01T12:34:56.789", res5.Str.ToString())

        Dim source6 = U8String.NewString("2023-10-01T12:34:56")
        Dim iter6 = source6.GetIterator()
        ' 解析を実行
        Dim res6 = TomlParser.ParseDateTime(source6, iter6)
        Assert.Equal(TomlExpressionType.LocalDateTime, res6.Type)

        Dim source7 = U8String.NewString("2023-10-01")
        Dim iter7 = source7.GetIterator()
        ' 解析を実行
        Dim res7 = TomlParser.ParseDateTime(source7, iter7)
        Assert.Equal(TomlExpressionType.LocalDate, res7.Type)

        Dim source8 = U8String.NewString("12:34:56")
        Dim iter8 = source8.GetIterator()
        ' 解析を実行
        Dim res8 = TomlParser.ParseDateTime(source8, iter8)
        Assert.Equal(TomlExpressionType.LocalTime, res8.Type)
        Assert.Equal("12:34:56", res8.Str.ToString())
    End Sub

    <Fact>
    Sub ParseMlBasicStringTest()
        Dim source = U8String.NewString("""""""Hello World" & vbCrLf & "This is a multi-line string.""""""")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseMlBasicString(source, iter)
        Assert.Equal(TomlExpressionType.MlBasicString, res.Type)
        Assert.Equal("""""""Hello World" & vbCrLf & "This is a multi-line string.""""""", res.Str.ToString())

        Dim errSource = U8String.NewString("""""""Hello World""")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseMlBasicString(errSource, errIter)
            End Sub
        )

        Dim source2 = U8String.NewString("""""""
Roses are red
Violets are blue""""""")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseMlBasicString(source2, iter2)
        Assert.Equal(TomlExpressionType.MlBasicString, res2.Type)
        Assert.Equal("""""""
Roses are red
Violets are blue""""""", res2.Str.ToString())

        Dim source3 = U8String.NewString("""""""\
       The quick brown \
       fox jumps over \
       the lazy dog.\
""""""")
        Dim iter3 = source3.GetIterator()
        ' 解析を実行
        Dim res3 = TomlParser.ParseMlBasicString(source3, iter3)
        Assert.Equal(TomlExpressionType.MlBasicString, res3.Type)
        Assert.Equal("""""""\
       The quick brown \
       fox jumps over \
       the lazy dog.\
""""""", res3.Str.ToString())

        Dim source4 = U8String.NewString("""""""""This,"" she said, ""is just a pointless statement.""""""""")
        Dim iter4 = source4.GetIterator()
        ' 解析を実行
        Dim res4 = TomlParser.ParseMlBasicString(source4, iter4)
        Assert.Equal(TomlExpressionType.MlBasicString, res4.Type)
        Assert.Equal("""""""""This,"" she said, ""is just a pointless statement.""""""""", res4.Str.ToString())
    End Sub

    <Fact>
    Sub ParseMlLiteralStringTest()
        Dim source = U8String.NewString("'''
Hello World
This is a multi-line string.
'''")
        Dim iter = source.GetIterator()
        ' 解析を実行
        Dim res = TomlParser.ParseMlLiteralString(source, iter)
        Assert.Equal(TomlExpressionType.MlLiteralString, res.Type)
        Assert.Equal("'''
Hello World
This is a multi-line string.
'''", res.Str.ToString())

        Dim errSource = U8String.NewString("'''
Hello World
")
        Dim errIter = errSource.GetIterator()
        ' 解析を実行
        Assert.Throws(Of TomlParseException)(
            Sub()
                TomlParser.ParseMlLiteralString(errSource, errIter)
            End Sub
        )

        Dim source2 = U8String.NewString("'''I [dw]on't need \d{2} apples'''")
        Dim iter2 = source2.GetIterator()
        ' 解析を実行
        Dim res2 = TomlParser.ParseMlLiteralString(source2, iter2)
        Assert.Equal(TomlExpressionType.MlLiteralString, res2.Type)
        Assert.Equal("'''I [dw]on't need \d{2} apples'''", res2.Str.ToString())

        Dim source3 = U8String.NewString("'''
生の文字列では、
最初の改行は取り除かれます。
    その他の空白は、
    保持されます。
'''")
        Dim iter3 = source3.GetIterator()
        ' 解析を実行
        Dim res3 = TomlParser.ParseMlLiteralString(source3, iter3)
        Assert.Equal(TomlExpressionType.MlLiteralString, res3.Type)
        Assert.Equal("'''
生の文字列では、
最初の改行は取り除かれます。
    その他の空白は、
    保持されます。
'''", res3.Str.ToString())
    End Sub

    <Fact>
    Sub ConvertToBasicStringTest()
        Dim src1 = U8String.NewString("""Hello World""")
        Dim iter1 = src1.GetIterator()
        Dim res1 = TomlParser.ConvertToBasicString(iter1)
        Assert.Equal("Hello World", res1.ToString())

        Dim src2 = U8String.NewString("""Hello\tWorld""")
        Dim iter2 = src2.GetIterator()
        Dim res2 = TomlParser.ConvertToBasicString(iter2)
        Assert.Equal($"Hello{vbTab}World", res2.ToString())

        Dim src3 = U8String.NewString("""Hello\nWorld""")
        Dim iter3 = src3.GetIterator()
        Dim res3 = TomlParser.ConvertToBasicString(iter3)
        Assert.Equal($"Hello{vbLf}World", res3.ToString())

        Dim src4 = U8String.NewString("""Hello\rWorld""")
        Dim iter4 = src4.GetIterator()
        Dim res4 = TomlParser.ConvertToBasicString(iter4)
        Assert.Equal($"Hello{vbCr}World", res4.ToString())

        Dim src5 = U8String.NewString("""あいう\bえお""")
        Dim iter5 = src5.GetIterator()
        Dim res5 = TomlParser.ConvertToBasicString(iter5)
        Assert.Equal($"あいう{vbBack}えお", res5.ToString())

        Dim src6 = U8String.NewString("""Hello \ud090 World""")
        Dim iter6 = src6.GetIterator()
        Dim res6 = TomlParser.ConvertToBasicString(iter6)
        Assert.Equal("Hello А World", res6.ToString())

        Dim src7 = U8String.NewString("""Hello \U00e29fb2 World""")
        Dim iter7 = src7.GetIterator()
        Dim res7 = TomlParser.ConvertToBasicString(iter7)
        Assert.Equal("Hello ⟲ World", res7.ToString())

        Dim src8 = U8String.NewString("""Hello\fWorld""")
        Dim iter8 = src8.GetIterator()
        Dim res8 = TomlParser.ConvertToBasicString(iter8)
        Assert.Equal($"Hello{vbFormFeed}World", res8.ToString())

        Dim src9 = U8String.NewString("""Hello\""World""")
        Dim iter9 = src9.GetIterator()
        Dim res9 = TomlParser.ConvertToBasicString(iter9)
        Assert.Equal("Hello""World", res9.ToString())

        Dim src10 = U8String.NewString("""Hello\\World""")
        Dim iter10 = src10.GetIterator()
        Dim res10 = TomlParser.ConvertToBasicString(iter10)
        Assert.Equal("Hello\World", res10.ToString())
    End Sub

    <Fact>
    Sub ConvertToMlBasicStringTest()
        Dim src1 = U8String.NewString("""""""
Roses are red
Violets are blue""""""")
        Dim iter1 = src1.GetIterator()
        Dim res1 = TomlParser.ConvertToMlBasicString(iter1)
        Assert.Equal("Roses are red" & vbCrLf & "Violets are blue", res1.ToString())
    End Sub

End Class
