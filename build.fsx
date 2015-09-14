// include FAKE libs
#r "packages/FAKE.4.4.1/tools/FakeLib.dll"

open Fake
open Fake.Testing

// directories
let testDir = "test/LastCellTests/bin/Debug/"

Target "xUnitTest" (fun _ ->
    !! (testDir @@ "*Tests.dll")
        |> xUnit2 (fun p -> {p with HtmlOutputPath = Some(testDir @@ ".xunit.html")})
)

RunTargetOrDefault "xUnitTest"