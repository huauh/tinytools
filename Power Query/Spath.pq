let
    Spath = (PathTable as text, Flag as text) as text =>
        Excel.CurrentWorkbook(){
            [Name = Text.From(PathTable)]
        }[Content]{
            [标识 = Text.From(Flag)]
        }[路径],

    Metadata = [
        Documentation.Name = "Spath",
        Documentation.Description = "返回数据源的绝对路径。需要先确保存在至少包含这两列( [标识] 和 [路径] )的智能表。PathTable: 表示智能表的名称；Flag: #(tab)表示路径所在行的[标识]列的值。",
        Documentation.Examples = {
            [
                Description = "从路径总表(PathTable)，获取产品表(Flag)的绝对路径",
                Code = "Spath(""路径总表"",""产品表"")",
                Result = "E:\目录1\目录2\目录3\目录4\资料链接\O5_产品信息表.xlsm"
            ]
        }
    ]
in
    Value.ReplaceType(
        Spath,
        Value.Type(
            Spath
        ) meta Metadata
    )