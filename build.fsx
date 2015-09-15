// include FAKE libs
#r "packages/FAKE.4.4.1/tools/FakeLib.dll"
#r "packages/FSharp.Formatting.2.10.3/lib/net40/FSharp.MetadataFormat.dll"
#r "packages/FSharp.Formatting.2.10.3/lib/net40/RazorEngine.dll"

open Fake
open Fake.Testing
open FSharp.MetadataFormat
open System.IO
open RazorEngine

// directories
let testDir = "test/LastCellTests/bin/Debug/"
let srcDir = "src/LastCell/bin/Debug/"
let outDir = "output/"
let tmpDir = "packages/FSharp.Formatting.2.10.3/templates"

Target "xUnitTest" (fun _ ->
    !! (testDir @@ "*Tests.dll")
        |> xUnit2 (fun p -> {p with HtmlOutputPath = Some(testDir @@ ".xunit.html")})
)

Target "dox" (fun _ -> 
    !! (srcDir @@ "*.dll")
        |> fun p -> MetadataFormat.Generate(srcDir @@ ".html", outDir, [ tmpDir ]) 
)

RunTargetOrDefault "xUnitTest"