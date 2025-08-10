Imports System
Imports Xunit
Imports ZoppaTomlLibrary.Strings
Imports ZoppaTomlLibrary.Toml

Public Class ExampleTest

    Private ReadOnly _doc As TomlDocument

    Public Sub New()
        Me._doc = TomlDocument.ReadFromFile("example.toml")
    End Sub

    <Fact>
    Public Sub ListKeyTest()
        Dim keys = Me._doc.GetKeys()
        Assert.NotEmpty(keys)
        Assert.Contains("title", keys)
        Assert.Contains("owner", keys)
        Assert.Contains("database", keys)
        Assert.Contains("servers", keys)
        Assert.Contains("clients", keys)
    End Sub

End Class
