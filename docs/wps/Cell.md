**Cell** 



代表单个表格单元格。Cell 对象是 [Cells](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Cells/Cells%20.htm#jsObject_Cells) 集合的成员。[Cells](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Cells/Cells%20.htm#jsObject_Cells) 集合代表指定对象中所有的单元格。

**说明**

若要返回 **Cell** 对象，应使用 [Table](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Table/Table%20.htm#jsObject_Table) 对象的**Cell**(*row*, *column*)（其中 *row* 是行号，*column* 是列号）；或使用 **Cells**(*index*)（其中 *index* 是索引编号）。以下示例对第一行的第二个单元格应用底纹。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let myCell = Application.ActiveDocument.Tables.Item(1).Cell(1,2) 	myCell.Shading.Texture = wdTexture20Percent }` |

以下示例对第一行的第一个单元格应用底纹。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Shading.Texture = wdTexture20Percent` |

使用**Add**方法将 **Cell** 对象添加到 **Cells** 集合。还可以使用 **Selection** 对象的 **InsertCells** 方法插入新的单元格。以下示例在 `myTable` 的第一个单元格之前添加一个单元格。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let myTable = Application.ActiveDocument.Tables.Item(1) 	myTable.Range.Cells.Add(myTable.Cell(1, 1)) }` |

以下示例设置引用第一个表格的头两个单元格的区域 (*myRange*)。设置此区域后，通过 **Merge** 方法合并单元格。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let myTable = Application.ActiveDocument.Tables.Item(1) 	let myRange = Application.ActiveDocument.Range(myTable.Cell(1, 1).Range.Start, myTable.Cell(1, 2).Range.End) 	myRange.Cells.Merge() }` |

可将**Add**方法与 **Rows** 或 **Columns** 集合配合使用，添加一行或一列单元格。

可将 **Information** 属性与 **Selection** 对象配合使用，返回当前行号和列号。以下示例更改选定内容中第一个单元格的宽度，然后显示该单元格的行号和列号。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	if(Application.Selection.Information(wdWithInTable) == true) {             Application.Selection.Cells.Item(1).Width = 22 	    alert("Cell " + Application.Selection.Information(wdStartOfRangeRowNumber) + "," + Application.Selection.Information(wdStartOfRangeColumnNumber)) 	} }` |

**方法**

|                                                              | 名称          | 说明                                                         |
| ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoSum**   | 插入一个 = (Formula) 域，以计算和显示表达式指定的位于表格某一单元格上方或左侧的各单元格的数值之和。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**    | 删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Formula**   | 将包含指定公式的 = (Formula) 域插入到单元格中。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge**     | 将指定单元格与另一表格单元格合并，成为一个单独的表格单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**    | 选择指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetHeight** | 设置表格单元格的高度。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetWidth**  | 设置表格列或单元格的宽度                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Split**     | 将一个表格单元格拆分为多个。                                 |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**            | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BottomPadding**      | 返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Column**             | 返回一个 **Column** 对象，该对象代表包含指定单元格的表格列。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnIndex**        | 返回包含指定单元格的表格列数。**Long** 类型，只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FitText**            | 如果 WPS 会减小键入单元格中的文字的可视大小，使其适应列宽，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**             | 返回或设置指定表格单元格的高度。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightRule**         | 返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格或行高度的规则。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                 | 在当前文档保存为网页时，返回或设置指定对象的识别标签。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftPadding**        | 返回或设置在表格的单个单元格或所有单元格的内容左侧的间距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NestingLevel**       | 返回指定单元格的嵌套层。**Long** 类型，只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**               | 返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的下一个表格单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回一个 **Object** 类型的值，该值代表指定 **Cell** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidth**     | 返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidthType** | 返回或设置用于指定单元格的宽度的首选度量单位。**WdPreferredWidthType** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**           | 返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的前一个表格单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**              | 返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定对象所含的部分文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightPadding**       | 返回或设置表格中某个单元格或全部单元格内容右侧的空格量（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Row**                | 返回一个 **Row** 对象，该对象代表指定单元格所在的行。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowIndex**           | 返回指定单元格所在行的行数。**Long** 类型，只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**            | 返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tables**             | 返回一个 **Tables** 集合，该集合代表指定表格单元格中的所有嵌套表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopPadding**         | 返回或设置表格中单独的单元格或全部单元格的内容上方要增加的间距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**  | 返回或设置表格中一个或多个单元格中文字的垂直对齐方式。**WdCellVerticalAlignment** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**              | 返回或设置指定表格单元格的宽度（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**           | 如果为 **True**，则 WPS 使文本自动换行并加长单元格，使单元格的宽度保持相同。**Boolean** 类型，可读写。 |

**成员方法**

#### **Cell.AutoSum**

插入一个 = (Formula) 域，以计算和显示表达式指定的位于表格某一单元格上方或左侧的各单元格的数值之和。

**语法**

**express.AutoSum()**

*express*   一个代表 **Cell** 对象的变量。

**说明**

有关 WPS 如何确定添加哪些值的信息，请参阅 **Formula** 方法。

**示例**

#### **Cell.Delete**

删除一个或多个表格单元格并可选择控制如何移动剩余的单元格。

**语法**

**express.Delete(ShiftCells)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *ShiftCells* | 可选          | **Variant**  | 剩余单元格移动的方向。可以是任意 WdDeleteCells 常量。如果省略，最后删除的单元格的右侧单元格向左移动。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例删除活动文档中第一个表格中的第一个单元格 */ Application.ActiveDocument.Tables.Item(1).Cell(1,1).Delete()` |

#### **Cell.Formula**

将包含指定公式的 = (Formula) 域插入到单元格中。

**语法**

**express.Formula(Formula, NumFormat)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Formula*   | 可选          | **Variant**  | 希望 = (Formula) 域求值的数学公式。可以使用与电子表格类似的方式引用表格中的单元格。例如，“=SUM(A4:C4)”指定第四行中的前三个值。有关 = (Formula) 域的详细内容，请参阅域代码：= (Formula) 域。 |
| *NumFormat* | 可选          | **Variant**  | = (Formula) 域的结果的格式。有关可应用的格式类型的信息，请参阅“数字图片 (\#)”域开关。 |

**说明**

如果插入点所在单元格的上面或左面至少有一个单元格包含数值，则 **Formula** 是可选的。如果插入点上面的单元格含有数值，则插入的域为 {=SUM(ABOVE)}；如果插入点左面的单元格含有数值，则插入的域为 {=SUM(LEFT)}；如果在插入点上面和左面的单元格中都含有数值，则 WPS 使用以下规则来确定插入哪一个 SUM 函数：

- 如果紧邻插入点上面的单元格中含有数值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果紧邻插入点上面的单元格中不包含数值，而紧邻插入点左面的单元格中含有数值，则 WPS 插入 {=SUM(LEFT)}。
- 如果紧邻单元格中均不包含数值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果未指定 **Formula**，而且插入点上面和左面的所有单元格均为空，则域的结果将会出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开始处创建一个 3x3 表格，然后计算第一列的平均值。*/ function test() { let myRange = ActiveDocument.Range(0, 0) let myTable = ActiveDocument.Tables.Add(myRange, 3, 3) myTable.Cell(1,1).Range.InsertAfter("100") myTable.Cell(2,1).Range.InsertAfter("50") myTable.Cell(3,1).Formula("=Average(Above)") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在插入点插入一个公式，该公式将决定选定单元格上面的单元格中的最大数值。*/ function test() { Selection.Collapse(wdCollapseStart) if(Selection.Information(wdWithInTable) == true) {     Selection.Cells.Item(1).Formula("=Max(Above)") } else {     alert("The insertion point is not in a table.") } }  ` |

#### **Cell.Merge**

将指定单元格与另一表格单元格合并，成为一个单独的表格单元格。

**语法**

**express.Merge(MergeTo)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**    | **说明**         |
| --------- | ------------- | --------------- | ---------------- |
| *MergeTo* | 必选          | **Cell object** | 要合并的单元格。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的表 1 的前两个单元格合并，然后删除表格的边框。*/ function test() {     if (Application.ActiveDocument.Tables.Count >= 1) {         let Tab = Application.ActiveDocument.Tables.Item(1)         Tab.Cell(1, 1).Merge(Tab.Cell(1, 2))         Tab.Borders.Enable = false     } }` |

#### **Cell.Select**

选择指定的对象。

**语法**

**express.Select()**

*express*   一个代表 **Cell** 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 处理 Selection 对象。

#### **Cell.SetHeight**

设置表格单元格的高度。

**语法**

**express.SetHeight(RowHeight, HeightRule)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**        | **说明**                       |
| ------------ | ------------- | ------------------- | ------------------------------ |
| *RowHeight*  | 必选          | **Variant**         | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选          | **WdRowHeightRule** | 用于确定指定单元格高度的方法。 |

**说明**

设置 **Cell** 对象的 **SetHeight** 属性可自动为整行设置该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定单元格的行高设置为不小于 18 磅。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.SetHeight(18, wdRowHeightAtLeast)     }     else {         alert("The insertion point is not in a table.")     } } ` |

#### **Cell.SetWidth**

设置表格列或单元格的宽度

**语法**

**express.SetWidth(ColumnWidth, RulerStyle)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**     | **说明**                        |
| ------------- | ------------- | ---------------- | ------------------------------- |
| *ColumnWidth* | 必选          | **Single**       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选          | **WdRulerStyle** | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。**WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一张表格，设置第二行第一个单元格宽度为 1.5 英寸。本示例保持表格中其他单元格的宽度。*/ function test() {     let newDoc = Application.Documents.Add()     let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)     myTable.Cell(2, 1).SetWidth(InchesToPoints(1.5), wdAdjustNone) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置包含插入点的单元格宽度为 36 磅。本示例缩小第一列的宽度以保持表格的右边界位置。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.Item(1).SetWidth(36, wdAdjustFirstColumn)     }     else {         alert("The insertion point is not in a table.")     } }` |

#### **Cell.Split**

将一个表格单元格拆分为多个。

**语法**

**express.Split(NumRows, NumColumns)**

*express*   一个代表 **Cell** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                             |
| ------------ | ------------- | ------------ | ------------------------------------ |
| *NumRows*    | 可选          | **Variant**  | 一个单元格或多个单元格拆分成的行数。 |
| *NumColumns* | 可选          | **Variant**  | 一个单元格或多个单元格拆分成的列数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将第一张表格的第一个单元格拆分为两个单元格 */ Application.ActiveDocument.Tables.Item(1).Cell(1, 1).Split(1, 2)` |

**成员属性**

#### **Cell.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Cell** 对象的变量。

#### **Cell.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Cell** 对象的变量。

#### **Cell.BottomPadding**

返回或设置在表格中单独的单元格或所有单元格内容的下方添加的间距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.BottomPadding**

*express*   一个代表 **Cell** 对象的变量。

**说明**

单独单元格的 **BottomPadding** 属性的设置值会覆盖整个表格的 **BottomPadding** 属性的设置值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个表格的底边距设为 40 像素。*/ Application.ActiveDocument.Tables.Item(1).BottomPadding = PixelsToPoints(40, true)` |

#### **Cell.Column**

返回一个 **Column** 对象，该对象代表包含指定单元格的表格列。只读。

**语法**

**express.Column**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个 3x5 表格并为偶数列设置底纹。*/ function test() {     Application.Selection.Collapse(wdCollapseStart)     let tableNew = Application.ActiveDocument.Tables.Add(Selection.Range, 3, 5)     for (let i = 1; i <= tableNew.Rows.Item(1).Cells.Count; i++) {         if ((tableNew.Rows.Item(1).Cells.Item(i).ColumnIndex) % 2 == 0) {             tableNew.Rows.Item(1).Cells.Item(i).Column.Shading.Texture = wdTexture10Percent         }     } }` |

#### **Cell.ColumnIndex**

返回包含指定单元格的表格列数。**Long** 类型，只读。

**语法**

**express.ColumnIndex**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建表格，选择第一行中的每个单元格，然后返回包含所选单元格的列数。*/ function test() {     let docNew = Application.Documents.Add()     let tableNew = docNew.Tables.Add(Selection.Range, 3, 3)     for (let i = 1; i <= tableNew.Rows.Item(1).Cells.Count; i++) {         tableNew.Rows.Item(1).Cells.Item(i).Select()         alert("This is column " + tableNew.Rows.Item(1).Cells.Item(i).ColumnIndex)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回插入点所在的单元格的列数。*/ alert(Application.Selection.Cells.Item(1).ColumnIndex)` |

#### **Cell.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Cell** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **Cell.FitText**

如果 WPS 会减小键入单元格中的文字的可视大小，使其适应列宽，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.FitText**

*express*   一个代表 **Cell** 对象的变量。

**说明**

如果将 **FitText** 设置为 **True**，则虽然字体尺寸没有改变，但将调整字符的可视宽度，使单元格可容纳所有键入的文字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定选定内容的第一单元格在其宽度内自动调整键入的文字大小。*/ Application.Selection.Cells.Item(1).FitText = true` |

#### **Cell.Height**

返回或设置指定表格单元格的高度。

**语法**

**express.Height**

*express*   一个代表 **Cell** 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto**，则 **Height** 返回 **wdUndefined**；如果设置 **Height** 属性，系统将 **HeightRule** 设置为 **wdRowHeightAtLeast**。可读/写 **Single** 类型。

#### **Cell.HeightRule**

返回或设置一个 **WdRowHeightRule** 常量，该常量代表确定指定单元格或行高度的规则。可读写。

**语法**

**express.HeightRule**

*express*   一个代表 **Cell** 对象的变量。

**说明**

设置 **Cell** 对象的 **HeightRule** 属性将自动设置整行的高度。

#### **Cell.Id**

在当前文档保存为网页时，返回或设置指定对象的识别标签。**String** 类型，可读写。

**语法**

**express.Id**

*express*   一个代表 **Cell** 对象的变量。

**说明**

可以将标签作为引用其他网页或当前文档中内容的超链接。

#### **Cell.LeftPadding**

返回或设置在表格的单个单元格或所有单元格的内容左侧的间距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.LeftPadding**

*express*   一个代表 **Cell** 对象的变量。

**说明**

单个单元格的 **LeftPadding** 属性设置将覆盖整个表格的 **LeftPadding** 属性的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个表格的第一个单元格的左边距设为 40 像素。*/ Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).LeftPadding = PixelsToPoints(40, false)` |

#### **Cell.NestingLevel**

返回指定单元格的嵌套层。**Long** 类型，只读。

**语法**

**express.NestingLevel**

*express*   一个代表 **Cell** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一个文档，创建一个三层嵌套表格，并在每个表格的第一个单元格中填入该表格所在的嵌套层数。*/ function test() {     Application.Documents.Add()     Application.ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)      let Rng1 = Application.ActiveDocument.Tables.Item(1).Range     Rng1.Copy()     Rng1.Cells.Item(1).Range.Text = Rng1.Cells.Item(1).NestingLevel     Rng1.Cells.Item(5).Range.PasteAsNestedTable()      let Rng2 = Rng1.Cells.Item(5).Tables.Item(1).Range     Rng2.Cells.Item(1).Range.Text = Rng2.Cells.Item(1).NestingLevel     Rng2.Cells.Item(5).Range.PasteAsNestedTable()      let Rng3 = Rng2.Cells.Item(5).Tables.Item(1).Range     Rng3.Cells.Item(1).Range.Text = Rng3.Cells.Item(1).NestingLevel }` |

#### **Cell.Next**

返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的下一个表格单元格。只读。

**语法**

**express.Next**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果选定内容位于表格之中，则本示例选定表格中下一个单元格的内容。 */ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.Item(1).Next.Select()     } } ` |

#### **Cell.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Cell** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中第一个表格的第一个单元格设置变量，将单元格的宽度更改为 36 磅，并删除表格的边框。*/ function test() {     let objCell = Application.ActiveDocument.Tables.Item(1).Cell(1, 1)     objCell.SetWidth(36, wdAdjustNone)     objCell.Parent.Borders.Enable = false }` |

#### **Cell.PreferredWidth**

返回或设置指定单元格的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。

**语法**

**express.PreferredWidth**

*express*   一个代表 **Cell** 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints**，则 **PreferredWidth** 属性返回或设置宽度（以磅为单位）。如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent**，则 **PreferredWidth** 属性返回或设置宽度（以窗口宽度的百分比表示）。

#### **Cell.PreferredWidthType**

返回或设置用于指定单元格的宽度的首选度量单位。**WdPreferredWidthType** 类型，只读。

**语法**

**express.PreferredWidthType**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为以窗口宽度的百分比来表示宽度，然后将文档中第一张表格的宽度设置为窗口宽度的 50%。*/ function test() {     let rng = Application.ActiveDocument.Tables.Item(1)     rng.PreferredWidthType = wdPreferredWidthPercent     rng.PreferredWidth = 50 } ` |

#### **Cell.Previous**

返回一个 **Cell** 对象，该对象代表 **Cells** 集合中的前一个表格单元格。只读。

**语法**

**express.Previous**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果选定内容位于表格中，则本示例选定前一单元格的内容。*/ function test() { if(Selection.Information(wdWithInTable) == true) {     Selection.Rows.Item(1).Cells.Item(3).Previous.Select() } }` |

#### **Cell.Range**

返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例复制第一个表格首行中第一个单元格中的内容 */ Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Range.Copy()` |

#### **Cell.RightPadding**

返回或设置表格中某个单元格或全部单元格内容右侧的空格量（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.RightPadding**

*express*   一个代表 **Cell** 对象的变量。

**说明**

某个单元格的 **RightPadding** 属性的设置将覆盖整个表格的 **RightPadding** 属性的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第一个表格首行中第一个单元格的右边距设置为 40 像素。*/ Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).RightPadding = PixelsToPoints(40, false)` |

#### **Cell.Row**

返回一个 **Row** 对象，该对象代表指定单元格所在的行。

**语法**

**express.Row**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为表格中插入点所在的行应用底纹。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         Application.Selection.Cells.Item(1).Row.Shading.Texture = wdTexture10Percent     }     else {         alert("The insertion point is not in a table.")     } }` |

#### **Cell.RowIndex**

返回指定单元格所在行的行数。**Long** 类型，只读。

**语法**

**express.RowIndex**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一个 3x3 表格，选定第一列的每个单元格，然后显示包含所选单元格的行的行号。*/ function test() {     let newDoc = Application.Documents.Add()     let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)     for (let i = 1; i <= myTable.Columns.Item(1).Cells.Count; i++) {         myTable.Columns.Item(1).Cells.Item(i).Select()         alert("This is row " + myTable.Columns.Item(1).Cells.Item(i).RowIndex)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示选定内容分第一行的行号。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         alert(Application.Selection.Cells.Item(1).RowIndex)     } }` |

#### **Cell.Shading**

返回一个 **Shading** 对象，该对象指明指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对第一个表格首行中的第一个单元格应用水平线底纹。*/ function test() {     if (Application.ActiveDocument.Tables.Count >= 1) {         Application.ActiveDocument.Tables.Item(1).Rows.Item(1).Cells.Item(1).Shading.Texture = wdTextureHorizontal     } }` |

#### **Cell.Tables**

返回一个 **Tables** 集合，该集合代表指定表格单元格中的所有嵌套表格。只读。

**语法**

**express.Tables**

*express*   一个代表 **Cell** 对象的变量。

#### **Cell.TopPadding**

返回或设置表格中单独的单元格或全部单元格的内容上方要增加的间距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.TopPadding**

*express*   一个代表 **Cell** 对象的变量。

**说明**

单独单元格的 **TopPadding** 属性设置将覆盖整个表格的 **TopPadding** 属性设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一张表格的顶部边距设为 40 像素。*/ Application.ActiveDocument.Tables.Item(1).TopPadding = PixelsToPoints(40, true)` |

#### **Cell.VerticalAlignment**

返回或设置表格中一个或多个单元格中文字的垂直对齐方式。**WdCellVerticalAlignment** 类型，可读写。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个新文档中创建一个 3x3 表格，并为表格中的每个单元格分配连续的单元格号。然后将第一行的高度设置为 20 磅，并在单元格的顶端垂直对齐文本。*/ function test() {     let newDoc = Application.Documents.Add()     let myTable = newDoc.Tables.Add(Selection.Range, 3, 3)     i = 1     for (let j = 1; j <= myTable.Range.Cells.Count; j++) {         myTable.Range.Cells.Item(j).Range.InsertAfter("cell" + i)         i++     }     let fpx = myTable.Rows.Item(1)     fpx.Height = 20     fpx.Cells.VerticalAlignment = wdAlignVerticalTop }` |

#### **Cell.Width**

返回或设置指定表格单元格的宽度（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.Width**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回插入点所在单元格的宽度（以英寸为单位）。*/ function test() {     if (Application.Selection.Information(wdWithInTable) == true) {         alert(Application.PointsToInches(Selection.Cells.Item(1).Width))     } }` |

#### **Cell.WordWrap**

如果为 **True**，则 WPS 使文本自动换行并加长单元格，使单元格的宽度保持相同。**Boolean** 类型，可读写。

**语法**

**express.WordWrap**

*express*   一个代表 **Cell** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 对第一张表格的第三个单元格中的文本自动换行，以使单元格的宽度保持相同。*/ Application.ActiveDocument.Tables.Item(1).Range.Cells.Item(3).WordWrap = true` |

适用环境：web

适用平台：windows/linux