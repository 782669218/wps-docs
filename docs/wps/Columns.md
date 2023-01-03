#### **Columns**



**Column** 对象的集合，这些对象代表表格中的列。

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**             | 返回一个 **Column** 对象，该对象代表添加到表中的列。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFit**         | 改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**          | 删除指定列。                                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DistributeWidth** | 将指定列调整为等宽。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**            | 返回集合中的单个 **Column** 对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**          | 选择指定表格列。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetWidth**        | 设置表格列的宽度。                                           |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**            | 返回一个 **Borders** 集合，该集合代表指定列的所有边框。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**              | 返回一个 **Number** 类型的值，该值代表集合中列的数目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **First**              | 返回一个 **Column** 对象，该对象代表在 **Columns** 集合中的第一个项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Last**               | 返回表示表格中最后一列的 **Column** 对象。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NestingLevel**       | 返回指定列的嵌套层。只读 **Number** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回一个 **Object** 类型值，该值代表指定 **Columns** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidth**     | 返回或设置指定列的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreferredWidthType** | 返回或设置用于指定单元格、列或表格的首选度量单位。可读/写 **WdPreferredWidthType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**            | 返回一个表示指定表格列的底纹格式的 **Shading** 对象。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**              | 返回或设置指定列的宽度，以磅为单位。可读/写 **Long** 类型。  |

**成员方法**

#### **Columns.Add**

返回一个 **Column** 对象，该对象代表添加到表中的列。

**语法**

**express.Add(BeforeColumn)**

*express*   一个代表 **Columns** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                   |
| -------------- | ------------- | ------------ | ------------------------------------------ |
| *BeforeColumn* | 可选          | **Variant**  | 代表将会直接显示在新列右侧的 Column 对象。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在活动文档中创建一个有两行两列的表格，然后在第一列之前添加一列。新列的宽度设为 1.5 英寸 */ function test() {     let myTable = Application.ActiveDocument.Tables.Add(Application.Selection.Range, 2, 2)     let newCol = myTable.Columns.Add(myTable.Columns.Item(1))     newCol.SetWidth(Application.InchesToPoints(1.5), wdAdjustNone) }` |

#### **Columns.AutoFit**

改变表格列宽，使之在单元格文本换行方式不变的情况下，适应文本宽度。

**语法**

**express.AutoFit()**

*express*   一个代表 **Columns** 对象的变量。

**说明**

如果表格的宽度已等于从左边界到右边界的距离，则此方法无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一个 3x3 表格，然后调整第一列的宽度，使之与文本的宽度相称。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 3, 3) tableNew.Cell(1,1).Range.InsertAfter("First cell") tableNew.Columns.Item(1).AutoFit() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一个 3x3 表格，然后调整所有列的宽度，使之与文本的宽度相称。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 3, 3) tableNew.Cell(1,1).Range.InsertAfter("First cell") tableNew.Cell(1,2).Range.InsertAfter("This is cell (1,2)") tableNew.Cell(1,3).Range.InsertAfter("(1,3)") tableNew.Columns.AutoFit() }` |

#### **Columns.Delete**

删除指定列。

**语法**

**express.Delete()**

*express*   一个代表 **Columns** 对象的变量。

#### **Columns.DistributeWidth**

将指定列调整为等宽。

**语法**

**express.DistributeWidth()**

*express*   一个代表 **Columns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个表格的各列调整为等宽。*/ Application.ActiveDocument.Tables.Item(1).Columns.DistributeWidth()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例调整选定单元格的宽度。*/ function test() {     if (Application.Selection.Cells.Count >= 2) {         Application.Selection.Cells.DistributeWidth()     } }` |

#### **Columns.Item**

返回集合中的单个 **Column** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Columns** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Number**   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

#### **Columns.Select**

选择指定表格列。

**语法**

**express.Select()**

*express*   一个代表 **Columns** 对象的变量。

**说明**

使用该方法后，可用 **Selection** 对象来处理选定项目。有关详细信息，请参阅 处理 Selection 对象。

#### **Columns.SetWidth**

设置表格列的宽度。

**语法**

**express.SetWidth(ColumnWidth, RulerStyle)**

*express*   一个代表 **Columns** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**     | **说明**                        |
| ------------- | ------------- | ---------------- | ------------------------------- |
| *ColumnWidth* | 必选          | **Single**       | 指定列的宽度，以磅为单位。      |
| *RulerStyle*  | 必选          | **WdRulerStyle** | 控制 WPS 调整单元格宽度的方式。 |

**说明**

上述 **WdRulerStyle** 行为应用于左对齐的表格。**WdRulerStyle** 行为应用于中对齐和右对齐的表格时可能会出现未知效果，因此应谨慎使用 **SetWidth** 方法。

**成员属性**

#### **Columns.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Columns** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Columns.Borders**

返回一个 **Borders** 集合，该集合代表指定列的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Columns** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Columns.Count**

返回一个 **Number** 类型的值，该值代表集合中列的数目。只读。

**语法**

**express.Count**

*express*   一个代表 **Columns** 对象的变量。

#### **Columns.First**

返回一个 **Column** 对象，该对象代表在 **Columns** 集合中的第一个项目。

**语法**

**express.First**

*express*   一个代表 **Columns** 对象的变量。

#### **Columns.Last**

返回表示表格中最后一列的 **Column** 对象。

**语法**

**express.Last**

*express*   一个代表 **Columns** 对象的变量。

#### **Columns.NestingLevel**

返回指定列的嵌套层。只读 **Number** 类型。

**语法**

**express.NestingLevel**

*express*   一个代表 **Columns** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

#### **Columns.Parent**

返回一个 **Object** 类型值，该值代表指定 **Columns** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Columns** 对象的变量。

#### **Columns.PreferredWidth**

返回或设置指定列的首选宽度（以磅为单位或表示为窗口宽度的百分比）。**Single** 类型，可读写。

**语法**

**express.PreferredWidth**

*express*   一个代表 **Columns** 对象的变量。

**说明**

如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPoints**，则 **PreferredWidth** 属性将返回或设置以磅为单位的宽度值；如果 **PreferredWidthType** 属性设置为 **wdPreferredWidthPercent**，则 **PreferredWidth** 属性将返回或设置用窗口宽度百分比表示的宽度值。

#### **Columns.PreferredWidthType**

返回或设置用于指定单元格、列或表格的首选度量单位。可读/写 **WdPreferredWidthType** 类型。

**语法**

**express.PreferredWidthType**

*express*   一个代表 **Columns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 以窗口宽度的百分比来接受宽度，然后将活动文档的第一个表格的所有列宽设置为窗口宽度的 50%。*/ function test() {     let rng = Application.ActiveDocument.Tables.Item(1).Columns     rng.PreferredWidthType = wdPreferredWidthPercent     rng.PreferredWidth = 50 }` |

#### **Columns.Shading**

返回一个表示指定表格列的底纹格式的 **Shading** 对象。

**语法**

**express.Shading**

*express*   一个代表 **Columns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档中第一个表格的所有列应用水平线纹理。*/ function test() {     if (Application.ActiveDocument.Tables.Count >= 1) {         Application.ActiveDocument.Tables.Item(1).Columns.Shading.Texture = wdTextureDiagonalDown     } }` |

#### **Columns.Width**

返回或设置指定列的宽度，以磅为单位。可读/写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **Columns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中创建一张 5x5 的表格，然后将该表格中的所有列宽设为 1.5 英寸。*/ function test() {     let objDoc = Application.ActiveDocument     let objTable = objDoc.Tables.Add(Selection.Range, 5, 5)     objTable.Columns.Width = InchesToPoints(1.5) }` |

适用环境：web

适用平台：windows/linux