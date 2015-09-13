#if INTERACTIVE
#r "Microsoft.Office.Interop.Excel"
#r "System.Runtime.InteropServices"
#r "Microsoft.Office.Core"
#r "xunit"
#r "Swensen.Unquote"
#else
module ComLastCellTests
#endif

open System
open ComLastCell
open Xunit
open Swensen.Unquote
open Microsoft.Office.Interop.Excel

let handler = ExcelHandler(ApplicationClass(Visible = true))

[<Fact>]
let ``some dumb test that should work`` () = 
    let actual = 256
    test <@ actual = 256 @>

[<Theory>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\c784b81f-7cd6-4a62-9fe2-34cdb799121b", 65536, 256, 341, 35)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\72710819-c760-4b4c-a99d-b2f7c7b2c529", 65536, 256, 1, 13)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\0c14e792-dfee-4316-acb3-056e461f3bfe", 65536, 256, 69, 11)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\4c144b78-967b-400d-b5cf-d315c07059e2", 65536, 256, 1, 14)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\6beedcdf-9410-41f3-a73e-b5ea8c4f1a3d", 65536, 256, 106, 47)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\16fe50f8-47a3-43a4-b11f-3b98fa46ff0d", 65536, 256, 8, 13)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\24e5ab7a-fae5-4383-b740-870a89f6528e", 65536, 256, 21, 13)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\072bae99-01be-4c68-8317-851d92ff7f25", 1048576, 16384, 96, 27)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\763ebe4b-dc71-45b1-9770-b284e1d5691e", 65536, 256, 3, 261)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\78902ae0-69a4-4545-92c6-1f967cfa84d1", 65536, 256, 69, 13)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\6795932d-d0a5-4225-a61a-b838d3faaf91", 65536, 256, 4, 10)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\a8182933-d859-4ddf-89f8-183ac766b282", 65536, 256, 118, 20)>]
[<InlineData(__SOURCE_DIRECTORY__ + "\\fusewb\\ebecc99b-3229-4fe9-bd54-c0f79ee9953d", 65536, 256, 405, 10)>]
let ``worksheetlastcell returns the correct result`` (filename : string) (maxr : int) (maxc : int) (lastr : int) (lastc : int) =
    let actual = examineFile handler filename
    test <@ (fst actual).MaxRow = maxr @>
    test <@ (fst actual).MaxCol = maxc @>
    test <@ (snd actual).LastRow = lastr @>
    test <@ (snd actual).LastCol = lastc @>
