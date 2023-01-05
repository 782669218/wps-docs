**Cell**



由表格列、表格行、选定内容或区域中的 **Cell**对象组成的集合。

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**              | 返回一个 **Cell** 对象，该对象代表添加到表格中的单元格。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFit**          | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DistributeHeight** | 将指定单元格调整为等高。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DistributeWidth**  | 将指定单元格调整为等宽。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**             | 返回集合中的单个 **Cell** 对象。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge**            | 将指定的多个表格单元格相互合并。合并后的结果是一个表格单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetHeight**        | 设定表格单元格的高度。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetWidth**         | 设置表格的列或单元格的宽度。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Split**            | 拆分表格单元格范围。                                         |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**            | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**              | 返回 **Cells** 集合中的项目数。**Number** 类型，只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**             | 返回或设置指定表格单元格的高度。可读/写 **Single** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightRule**         | 返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格高度的规则。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NestingLevel**       | 返回指定单元格的嵌套层。**Long** 类型，只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回一个 **Object** 类型的值，该值代表指定 **Cells** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidth**     | 返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidthType** | 返回或设置用于指定单元格的宽度的首选度量单位。**WdPreferredWidthType** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**            | 返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**  | 返回或设置表格中一个或多个单元格中文字的垂直对齐方式。**WdCellVerticalAlignment** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**              | 返回或设置表格单元格的宽度（以磅为单位）。**Number** 类型，可读写。 |

**成员方法**

#### **Cells.Add**

返回一个 **Cell** 对象，该对象代表添加到表格中的单元格。

**语法**

**express.Add(BeforeCell)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *BeforeCell* | 可选          | **Variant**  | 指定一个 Cell 对象，该对象代表将紧靠新单元格右侧显示的单元格。 |

**返回值**

Cell

#### **Cells.AutoFit**

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit()**

*express*   一个代表 **Cells** 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 3, 3) tableNew.Cell(1,1).Range.InsertAfter("First cell") tableNew.Cell(1,2).Range.InsertAfter("This is cell (1,2)") tableNew.Cell(1,3).Range.InsertAfter("(1,3)") tableNew.Columns.AutoFit() }` |

#### **Cells.Delete**

删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。

**语法**

**express.Delete(ShiftCells)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *ShiftCells* | 可选          | **Variant**  | 剩余单元格移动的方向。可以是任意 WdDeleteCells 常量。如果省略，最后删除的单元格的右侧单元格向左移动。 |

#### **Cells.DistributeHeight**

将指定单元格调整为等高。

**语法**

**express.DistributeHeight()**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的第一张表格的高度调整为相等。*/ Application.ActiveDocument.Tables.Item(1).Rows.DistributeHeight()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一张表格的前三行高度调整为相等。*/ function test() {     if (Application.ActiveDocument.Tables.Count >= 1) {         let rngTemp = Application.ActiveDocument.Range(Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Start, Application.ActiveDocument.Tables.Item(1).Rows.Item(3).Range.End)         rngTemp.Rows.DistributeHeight()     } }` |

#### **Cells.DistributeWidth**

将指定单元格调整为等宽。

**语法**

**express.DistributeWidth()**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 调整选中单元格，使宽度一致 */ Application.Selection.Cells.DistributeWidth() ` |

#### **Cells.Item**

返回集合中的单个 **Cell** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**返回值**

Cell

#### **Cells.Merge**

将指定的多个表格单元格相互合并。合并后的结果是一个表格单元格。

**语法**

**express.Merge()**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容第一行中的多个单元格合并为一个单元格，然后为该行应用底纹。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         let myrow = Application.Selection.Rows.Item(1)         myrow.Cells.Merge()         myrow.Shading.Texture = wdTexture10Percent     } } ` |

#### **Cells.SetHeight**

设定表格单元格的高度。

**语法**

**express.SetHeight(RowHeight, HeightRule)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**        | **说明**                       |
| ------------ | ------------- | ------------------- | ------------------------------ |
| *RowHeight*  | 必选          | **Variant**         | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选          | **WdRowHeightRule** | 确定指定单元格高度的方法。     |

**说明**

设置 **Cells** 对象的 **SetHeight** 属性将自动设置整行的高度。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选单元格的行高设置为至少 18 磅。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.SetHeight(18, wdRowHeightAtLeast)     }     else {         alert("The insertion point is not in a table.")     } }` |

#### **Cells.SetWidth**

设置表格的列或单元格的宽度。

**语法**

**express.SetWidth(ColumnWidth, RulerStyle)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**     | **说明**                        |
| ------------- | ------------- | ---------------- | ------------------------------- |
| *ColumnWidth* | 必选          | **Single**       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选          | **WdRulerStyle** | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为适用于左对齐的表格。**WdRulerStyle** 行为用于居中和右对齐的表格时可能无法预测，因此 **SetWidth** 方法应谨慎使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中创建一张表格，并将第二行第一个单元格的宽度设置为 1.5 英寸。此示例保持表格中其他单元格的宽度。*/ function test() {     let newDoc = Application.Documents.Add()     let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)     myTable.Cell(2, 1).SetWidth(InchesToPoints(1.5), wdAdjustNone) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将包含插入点的单元格的宽度设置为 36 磅。此示例还缩小第一列的宽度以保持表格的右边缘位置。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.Item(1).SetWidth(36, wdAdjustFirstColumn)     }     else {         alert("The insertion point is not in a table.")     } }` |

#### **Cells.Split**

拆分表格单元格范围。

**语法**

**express.Split(NumRows, NumColumns, MergeBeforeSplit)**

*express*   一个代表 **Cells** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                    |
| ------------------ | ------------- | ------------ | ----------------------------------------------------------- |
| *NumRows*          | 可选          | **Variant**  | 单元格或单元格组拆分成的行数。                              |
| *NumColumns*       | 可选          | **Variant**  | 单元格或单元格组拆分成的列数。                              |
| *MergeBeforeSplit* | 可选          | **Variant**  | 如果该参数值为 True，则在拆分多个单元格前会将它们相互合并。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例首先将选定单元格合并为一个单元格，然后将该单元格拆分为同一行中的三个单元格。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.Split(1, 3, true)     } }` |

**成员属性**

#### **Cells.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Cells** 对象的变量。

#### **Cells.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Cells** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档中第一个表格首行中的单元格应用内部和外部边框。*/ function test() {     let objTable = Application.ActiveDocument.Tables.Item(1)     let Bor = objTable.Rows.Item(1).Cells.Borders     Bor.InsideLineStyle = wdLineStyleSingle     Bor.OutsideLineStyle = wdLineStyleDouble }` |

#### **Cells.Count**

返回 **Cells** 集合中的项目数。**Number** 类型，只读。

**语法**

**express.Count**

*express*   一个代表 **Cells** 对象的变量。

#### **Cells.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Cells** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **Cells.Height**

返回或设置指定表格单元格的高度。可读/写 **Single** 类型。

**语法**

**express.Height**

*express*   一个代表 **Cells** 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto**，则 **Height** 返回 **wdUndefined**；如果设置 **Height** 属性，系统会将 **HeightRule** 设置为 **wdRowHeightAtLeast**。

#### **Cells.HeightRule**

返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格高度的规则。可读写。

**语法**

**express.HeightRule**

*express*   一个代表 **Cells** 对象的变量。

**说明**

设置 **Cells** 集合的 **HeightRule** 属性将自动设置整行的高度。

#### **Cells.NestingLevel**

返回指定单元格的嵌套层。**Long** 类型，只读。

**语法**

**express.NestingLevel**

*express*   一个代表 **Cells** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一个文档，创建一个三层嵌套表格，并在每个表格的第一个单元格中填入该表格所在的嵌套层数。*/ function test() {     Application.Documents.Add()     Application.ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)     let Rng1 = Application.ActiveDocument.Tables.Item(1).Range     Rng1.Copy()     Rng1.Cells.Item(1).Range.Text = Rng1.Cells.NestingLevel     Rng1.Cells.Item(5).Range.PasteAsNestedTable()     let Rng2 = Rng1.Cells.Item(5).Tables.Item(1).Range     Rng2.Cells.Item(1).Range.Text = Rng2.Cells.NestingLevel     Rng2.Cells.Item(5).Range.PasteAsNestedTable()     let Rng3 = Rng2.Cells.Item(5).Tables.Item(1).Range     Rng3.Cells.Item(1).Range.Text = Rng3.Cells.NestingLevel }` |

#### **Cells.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Cells** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Cells** 对象的变量。

#### **Cells.PreferredWidth**

返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。

**语法**

**express.PreferredWidth**

*express*   一个代表 **Cells** 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints**，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent**，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

#### **Cells.PreferredWidthType**

返回或设置用于指定单元格的宽度的首选度量单位。**WdPreferredWidthType** 类型，只读。

**语法**

**express.PreferredWidthType**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为以窗口宽度的百分比来表示宽度，然后将文档中第一张表格的宽度设置为窗口宽度的 50%。*/ function test() {     let rng = Application.ActiveDocument.Tables.Item(1)     rng.PreferredWidthType = wdPreferredWidthPercent     rng.PreferredWidth = 50 }` |

#### **Cells.Shading**

返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Cells** 对象的变量。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对表格 1 的第一行应用水平线纹理。*/function test() {    if (Application.ActiveDocument.Tables.Count >= 1) {        Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Shading.Texture = wdTextureHorizontal    }} ` |

#### **Cells.VerticalAlignment**

返回或设置表格中一个或多个单元格中文字的垂直对齐方式。**WdCellVerticalAlignment** 类型，可读写。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个新文档中创建一个 3x3 表格，并为表格中的每个单元格分配连续的单元格号。然后将第一行的高度设置为 20 磅，并在单元格的顶端垂直对齐文本。*/ function test() {     let newDoc = Application.Documents.Add()     let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)     let i = 1     for (let j = 1; j <= myTable.Range.Cells.Count; j++) {         myTable.Range.Cells.Item(j).Range.InsertAfter("Cell " + j)         i++     }     let rng = myTable.Rows.Item(1)     rng.Height = 20     rng.Cells.VerticalAlignment = wdAlignVerticalTop }` |

#### **Cells.Width**

返回或设置表格单元格的宽度（以磅为单位）。**Number** 类型，可读写。

**语法**

**express.Width**

*express*   一个代表 **Cells** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取选区单元格的宽度 */ alert(Application.Selection.Cells.Width);` |

适用环境：web

适用平台：windows/linux