Imports System
Imports ZoppaTomlLibrary.Toml

Module Program
    Sub Main(args As String())
        Dim doc = TomlDocument.ReadFromFile("example.toml")

        Console.WriteLine("title : {0}", doc("title").ValueTo(Of String)())

        Console.WriteLine("owner.name : {0}", doc("owner")("name").ValueTo(Of String)())
        Console.WriteLine("owner.dob : {0}", doc("owner")("dob").ValueTo(Of DateTimeOffset)())

        Console.WriteLine("database.server : {0}", doc("database.server").ValueTo(Of String)())
        Console.WriteLine("database.ports[0] : {0}", doc("database.ports")(0).ValueTo(Of Integer)())
        Console.WriteLine("database.ports[1] : {0}", doc("database.ports")(1).ValueTo(Of Integer)())
        Console.WriteLine("database.ports[2] : {0}", doc("database.ports")(2).ValueTo(Of Integer)())
        Console.WriteLine("database.connection_max : {0}", doc("database.connection_max").ValueTo(Of Integer)())
        Console.WriteLine("database.enabled : {0}", doc("database.enabled").ValueTo(Of Boolean)())

        Dim servers = doc("servers")
        Console.WriteLine("servers.alpha.ip : {0}", servers("alpha")("ip").ValueTo(Of String)())
        Console.WriteLine("servers.alpha.dc : {0}", servers("alpha")("dc").ValueTo(Of String)())
        Console.WriteLine("servers.beta.ip : {0}", servers("beta")("ip").ValueTo(Of String)())
        Console.WriteLine("servers.beta.dc : {0}", servers("beta")("dc").ValueTo(Of String)())

        Console.WriteLine("clients.data[0][0] : {0}", doc("clients")("data")(0)(0).ValueTo(Of String)())
        Console.WriteLine("clients.data[0][1] : {0}", doc("clients")("data")(0)(1).ValueTo(Of String)())
        Console.WriteLine("clients.data[1][0] : {0}", doc("clients")("data")(1)(0).ValueTo(Of Integer)())
        Console.WriteLine("clients.data[1][1] : {0}", doc("clients")("data")(1)(1).ValueTo(Of Integer)())

        Console.WriteLine("clients.hosts[0] : {0}", doc("clients")("hosts")(0).ValueTo(Of String)())
        Console.WriteLine("clients.hosts[1] : {0}", doc("clients")("hosts")(1).ValueTo(Of String)())
    End Sub
End Module
