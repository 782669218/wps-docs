#### **Rows**



**Row** 对象的集合，这些对象代表指定的选定内容、区域或表格中的表格行

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**              | 返回一个 **Row** 对象，该对象代表添加到表中的行。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToText**    | 将表格中的行转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 删除指定的表格行                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DistributeHeight** | 将指定行或单元格的高度调整为相等。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**             | 返回集合中的单个 **Row** 对象。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**           | 选择表格中行的集合。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetHeight**        | 设置表格行的高度。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetLeftIndent**    | 设置表格中一行或多行的缩进。                                 |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment**                  | 返回或设置一个 **WdRowAlignment** 常量，该常量代表指定行的对齐方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowBreakAcrossPage**       | 如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowOverlap**               | 返回或设置一个值，该值指定是否允许指定行与其他行重叠。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**                    | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                      | 返回一个 **Long** 类型的值，该值代表集合中的行数。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceBottom**             | 返回或设置文档文字与指定表格下边缘之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceLeft**               | 返回或设置文档文字与指定表格的左边缘之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceRight**              | 返回或设置文档文字与指定表格右边缘之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DistanceTop**                | 返回或设置文档文字与指定表格的上边缘之间的距离（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **First**                      | 返回一个 **Row** 对象，该对象代表 **Rows** 集合中的第一个项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeadingFormat**              | 如果将指定一行或数行的格式设置为表格标题，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                     | 返回或设置表格中指定的行的高度。Single 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeightRule**                 | 返回或设置确定指定单元格或行高度的规则。**WdRowHeightRule** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalPosition**         | 返回或设置由 **RelativeHorizontalPosition** 属性指定的项目与行边缘之间的水平距离。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Last**                       | 将 **Rows** 集合中的最后一项作为 **Row** 对象返回。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftIndent**                 | 返回或设置一个 **Single** 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NestingLevel**               | 返回指定表格行的嵌套层。只读 **Long** 类型。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回一个 **Object** 类型值，该值代表指定 **Rows** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeHorizontalPosition** | 指定一组行的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RelativeVerticalPosition**   | 指定一组行的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                    | 返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBetweenColumns**        | 返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TableDirection**             | 返回或设置 WPS 对指定的表格或行中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalPosition**           | 返回或设置由 **RelativeVerticalPosition** 属性指定的项目与行边缘之间的垂直距离。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WrapAroundText**             | 返回或设置文本是否环绕指定行。**Long** 类型，可读写。        |

**成员方法**

#### **Rows.Add**

返回一个 **Row** 对象，该对象代表添加到表中的行。

**语法**

**express.Add()**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在选定内容的第一行之前插入一个新行。 function test() {     if (Application.Selection.Information(wdWithInTable))         Application.Selection.Rows.Add(Selection.Rows.Item(1)) }  //本示例在第一张表中添加一行，然后将文本 Cell 插入该行。 function test() {	     let tblNew = Application.ActiveDocument.Tables.Item(1)     let rowNew = tblNew.Rows.Add(tblNew.Rows.Item(1))     for(i = 1; i <= rowNew.Cells.Count; ++i)     {     	let celTable = rowNew.Cells.Item(i)         celTable.Range.InsertAfter("Cell " + i)     } } ` |

#### **Rows.ConvertToText**

将表格中的行转换为文本并返回一个 **Range** 对象，该对象代表带分隔符的文本。

**语法**

**express.ConvertToText(Separator, NestedTables)**

*express*   一个代表 **Rows** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Separator*    | 可选          | **Variant**  | 用以分隔被转换列的字符（被转换行由段落标记分隔）。可以是下列任何 WdTableFieldSeparator 常量：wdSeparateByCommas、wdSeparateByDefaultListSeparator、wdSeparateByParagraphs 或 wdSeparateByTabs（默认值）。 |
| *NestedTables* | 可选          | **Variant**  | 如果要将嵌套的表格转换为文本，则为 True。如果 Separator 不是 wdSeparateByParagraphs，则将忽略此参数。默认值为 True。 |

#### **Rows.Delete**

删除指定的表格行

**语法**

**express.Delete()**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//删除选中区域的表格行 Application.Selection.Rows.Delete()` |

#### **Rows.DistributeHeight**

将指定行或单元格的高度调整为相等。

**语法**

**express.DistributeHeight()**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的第一张表格的高度调整为相等。*/ ActiveDocument.Tables.Item(1).Rows.DistributeHeight()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一张表格的前三行高度调整为相等。*/ function test() { if(ActiveDocument.Tables.Count >= 1) {     let rngTemp = ActiveDocument.Range(ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Start, ActiveDocument.Tables.Item(1).Rows.Item(3).Range.End)     rngTemp.Rows.DistributeHeight() } }` |

#### **Rows.Item**

返回集合中的单个 **Row** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Rows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

**返回值**

Row

#### **Rows.Select**

选择表格中行的集合。

**语法**

**express.Select()**

*express*   一个代表 **Rows** 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

#### **Rows.SetHeight**

设置表格行的高度。

**语法**

**express.SetHeight(RowHeight, HeightRule)**

*express*   一个代表 **Rows** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**        | **说明**                       |
| ------------ | ------------- | ------------------- | ------------------------------ |
| *RowHeight*  | 必选          | **Single**          | 一行或多行的高度，以磅为单位。 |
| *HeightRule* | 必选          | **WdRowHeightRule** | 用于确定指定行的高度的规则。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例创建一个表格，然后将表格中所有行的行高设置为 0.5 英寸（36 磅）。*/ function test() { let newDoc = Documents.Add() let aTable = newDoc.Tables.Add(Selection.Range, 3, 3) aTable.Rows.SetHeight(InchesToPoints(0.5), wdRowHeightExactly) }  ` |

#### **Rows.SetLeftIndent**

设置表格中一行或多行的缩进。

**语法**

**express.SetLeftIndent(LeftIndent, RulerStyle)**

*express*   一个代表 **Rows** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**     | **说明**                                                     |
| ------------ | ------------- | ---------------- | ------------------------------------------------------------ |
| *LeftIndent* | 必选          | **Single**       | 指定的一行或多行的当前左边缘与所需左边缘之间的距离（以磅为单位）。 |
| *RulerStyle* | 必选          | **WdRulerStyle** | 控制当左缩进发生变化时 WPS 调整表格的方式。WdRulerStyle 行为适用于左对齐的表格。居中和右对齐表格的 WdRulerStyle 行为不可预测；在这种情况下，请谨慎使用 SetLeftIndent 方法。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中创建一个表格，并将表格中的所有行缩进 0.5 英寸（36 磅）。在更改左缩进时，单元格宽度将调整以保持表格的右边缘。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 3, 3) tableNew.Rows.SetLeftIndent(InchesToPoints(0.5), wdAdjustSameWidth) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中“表格 1”的第一行缩进 18 磅，并通过缩小第一列的宽度来保持表格的右边界位置。*/ function test() { if(ActiveDocument.Tables.Count >= 1) {     ActiveDocument.Tables.Item(1).Rows.SetLeftIndent(18, wdAdjustFirstColumn) } }  ` |

**成员属性**

#### **Rows.Alignment**

返回或设置一个 **WdRowAlignment** 常量，该常量代表指定行的对齐方式。可读写。

**语法**

**express.Alignment**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一张表格的所有行都设置为居中对齐。*/ function CenterRows() {     ActiveDocument.Tables.Item(1).Rows.Alignment = wdAlignRowCenter }` |

#### **Rows.AllowBreakAcrossPage**

如果允许分页符拆分表格中一行或多行中的文本，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.AllowBreakAcrossPage**

*express*   一个代表 **Rows** 对象的变量。

**说明**

该属性可以是 **True**、**False** 或 **wdUndefined**（仅允许拆分部分指定文本）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一篇包含 5x5 表格的文档，并防止在分页时拆分表格的第三行。*/ function test() { let docNew = Documents.Add() let tableNew = docNew.Tables.Add(Selection.Range, 5, 5)  tableNew.Rows.Item(3).AllowBreakAcrossPages = false }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定分页时是否可以拆分当前表格中的行。如果插入点不在表格中，则显示一个消息框。*/ function test() { Selection.Collapse(wdCollapseStart)  if(!Selection.Tables.Count) {     MsgBox("The insertion point is not in a table.") } else {     let lngAllowBreak = Selection.Rows.AllowBreakAcrossPages } }` |

#### **Rows.AllowOverlap**

返回或设置一个值，该值指定是否允许指定行与其他行重叠。

**语法**

**express.AllowOverlap**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果指定行包含重叠行和非重叠行，则该属性返回 **wdUndefined**。可以设置为 **True** 或 **False**。**Long** 类型，可读写。如果将 **AllowOverlap** 设置为 **True**，会将 **WrapAroundText** 也设置为 **True**；如果将 **WrapAroundText** 设置为 **False**，则会将 **AllowOverlap** 也设置为 **False**。

因为 HTML 不支持表格或图形的重叠，所以在 Web 版式视图中将忽略 **AllowOverlap**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定文字环绕选定的表格，并且此表不能与其他被文字环绕的表格重叠。*/ function test() { Selection.Rows.WrapAroundText = true Selection.Rows.AllowOverlap = false }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定活动文档中的第一个图形可与其他图形重叠。*/ ActiveDocument.Shapes.Item(1).WrapFormat.AllowOverlap = true` |

#### **Rows.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Rows** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Rows.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Rows** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test() { let myTable = ActiveDocument.Tables.Item(1) let myBorders = myTable.Rows.Borders myBorders.InsideLineStyle = wdLineStyleSingle myBorders.OutsideLineStyle = wdLineStyleDouble }` |

#### **Rows.Count**

返回一个 **Long** 类型的值，该值代表集合中的行数。只读。

**语法**

**express.Count**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//获取激活文档的表格1的行数 Application.ActiveDocument.Tables.Item(1).Rows.Count` |

#### **Rows.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Rows.DistanceBottom**

返回或设置文档文字与指定表格下边缘之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.DistanceBottom**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False**，则该属性将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/ function test() { let myRows = ActiveDocument.Tables.Item(1).Rows myRows.WrapAroundText = true myRows.DistanceLeft = 20 myRows.DistanceRight = 20 myRows.DistanceTop = 20 myRows.DistanceBottom = 20 }` |

#### **Rows.DistanceLeft**

返回或设置文档文字与指定表格的左边缘之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.DistanceLeft**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False**，则该属性将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/ function test() { let myRows = ActiveDocument.Tables.Item(1).Rows myRows.WrapAroundText = true myRows.DistanceLeft = 20 myRows.DistanceRight = 20 myRows.DistanceTop = 20 myRows.DistanceBottom = 20 }` |

#### **Rows.DistanceRight**

返回或设置文档文字与指定表格右边缘之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.DistanceRight**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False**，则该属性将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/ function test() { let myRows = ActiveDocument.Tables.Item(1).Rows myRows.WrapAroundText = true myRows.DistanceLeft = 20 myRows.DistanceRight = 20 myRows.DistanceTop = 20 myRows.DistanceBottom = 20 }` |

#### **Rows.DistanceTop**

返回或设置文档文字与指定表格的上边缘之间的距离（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.DistanceTop**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果 **WrapAroundText** 为 **False**，则该属性将不起作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定文字环绕在活动文档中第一个表格的周围，并且环绕文字与该表格各边的间距为 20 磅。*/ function test() { let myRows = ActiveDocument.Tables.Item(1).Rows myRows.WrapAroundText = true myRows.DistanceLeft = 20 myRows.DistanceRight = 20 myRows.DistanceTop = 20 myRows.DistanceBottom = 20 }` |

#### **Rows.First**

返回一个 **Row** 对象，该对象代表 **Rows** 集合中的第一个项目。

**语法**

**express.First**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档第一个表格的第一行应用底纹和下边框格式。*/ function test() { ActiveDocument.Tables.Item(1).Borders.Enable = false let myFirst = ActiveDocument.Tables.Item(1).Rows.First myFirst.Shading.Texture = wdTexture10Percent myFirst.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle }` |

#### **Rows.HeadingFormat**

如果将指定一行或数行的格式设置为表格标题，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.HeadingFormat**

*express*   一个代表 **Rows** 对象的变量。

**说明**

当表格不止一页时，可将多行的格式重复设置为表格标题。属性值可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开始处创建一个 5x5 表格，然后将表格第一行的格式设置为表格标题。*/ function test() { let rngTemp = ActiveDocument.Range(0, 0) let tableNew = ActiveDocument.Tables.Add(rngTemp, 5, 5) tableNew.Rows.HeadingFormat = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判定是否将插入点所在行的格式设置为表格标题。*/ function test() { if(Selection.Information(wdWithInTable)) {     if(Selection.Rows.HeadingFormat) {         MsgBox("The current row is a table heading")     } } else {     MsgBox("The insertion point is not in a table.") } }` |

#### **Rows.Height**

返回或设置表格中指定的行的高度。Single 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果指定行的 **HeightRule** 属性为 **wdRowHeightAuto**，则 **Height** 返回 **wdUndefined**；设置 **Height** 属性会将 **HeightRule** 设置为 **wdRowHeightAtLeast**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第一张表格的行高度设置为至少 20 磅。*/ ActiveDocument.Tables.Item(1).Rows.Height = 20` |

#### **Rows.HeightRule**

返回或设置确定指定单元格或行高度的规则。**WdRowHeightRule** 类型，可读写。

**语法**

**express.HeightRule**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将确定所选行高度的规则设为按行中最高的单元格自动调整。*/ function test() { if(Selection.Information(wdWithInTable)) {     Selection.Rows.HeightRule = wdRowHeightAuto } else {     MsgBox("The insertion point is not in a table.") } }` |

#### **Rows.HorizontalPosition**

返回或设置由 **RelativeHorizontalPosition** 属性指定的项目与行边缘之间的水平距离。可读写 **Single** 类型。

**语法**

**express.HorizontalPosition**

*express*   一个代表 **Rows** 对象的变量。

**说明**

该属性可以是一个以磅为单位表示度量的数字，也可以是 **WdTablePosition** 常量之一。如果 **WrapAroundText** 属性为 **False**，则该属性无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个表格与右边距水平对齐。*/ function test() { if(ActiveDocument.Tables.Count >= 1) {     let myRows = ActiveDocument.Tables.Item(1).Rows     myRows.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin     myRows.HorizontalPosition = wdTableRight } }` |

#### **Rows.Last**

将 **Rows** 集合中的最后一项作为 **Row** 对象返回。

**语法**

**express.Last**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例删除活动文档中第一个表格中的最后一行。*/ ActiveDocument.Tables.Item(1).Rows.Last.Cells.Delete()` |

#### **Rows.LeftIndent**

返回或设置一个 **Single** 类型的值，该值代表指定表格行的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中第一个表格的各行设置左缩进。*/ ActiveDocument.Tables.Item(1).Rows.LeftIndent = InchesToPoints(1)` |

#### **Rows.NestingLevel**

返回指定表格行的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

*express*   一个代表 **Rows** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一个文档，创建一个三层嵌套表格，并在每张表格的第一个单元格中填入该表格所在的嵌套层数。*/ function test() { Documents.Add() ActiveDocument.Tables.Add(Selection.Range, 3, 3, wdWord9TableBehavior, wdAutoFitContent)  let myRange1 = ActiveDocument.Tables.Item(1).Range myRange1.Copy() myRange1.Cells.Item(1).Range.Text = myRange1.Cells.Item(1).NestingLevel myRange1.Cells.Item(5).Range.PasteAsNestedTable() 	 let myRange2 = myRange1.Cells.Item(5).Tables.Item(1).Range myRange2.Cells.Item(1).Range.Text = myRange2.Cells.Item(1).NestingLevel myRange2.Cells.Item(5).Range.PasteAsNestedTable()  let myRange3 = myRange2.Cells.Item(5).Tables.Item(1).Range myRange3.Cells.Item(1).Range.Text = myRange3.Cells.Item(1).NestingLevel }` |

#### **Rows.Parent**

返回一个 **Object** 类型值，该值代表指定 **Rows** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Rows** 对象的变量。

#### **Rows.RelativeHorizontalPosition**

指定一组行的相对水平位置。可读写 **WdRelativeHorizontalPosition** 类型。

**语法**

**express.RelativeHorizontalPosition**

*express*   一个代表 **Rows** 对象的变量。

#### **Rows.RelativeVerticalPosition**

指定一组行的相对垂直位置。可读写 **WdRelativeVerticalPosition** 类型。

**语法**

**express.RelativeVerticalPosition**

*express*   一个代表 **Rows** 对象的变量。

#### **Rows.Shading**

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容的第一段应用黄色底纹。*/ function test() { let myShading = Selection.Paragraphs.Item(1).Shading myShading.Texture = wdTexture12Pt5Percent myShading.BackgroundPatternColorIndex = wdYellow myShading.ForegroundPatternColorIndex = wdBlack }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将横线纹理应用于表格 1 的第一行。*/ function test() { if(ActiveDocument.Tables.Count >= 1) {     let myShading = ActiveDocument.Tables.Item(1).Rows.Item(1).Shading     myShading.Texture = wdTextureHorizontal } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一个单词的底纹设置为 10%。*/ ActiveDocument.Words.Item(1).Shading.Texture = wdTexture10Percent` |

#### **Rows.SpaceBetweenColumns**

返回或设置指定的一行或多行中相邻列的文本之间的距离（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.SpaceBetweenColumns**

*express*   一个代表 **Rows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回所选表格行中各列之间的距离（以磅为单位）。*/ function test() { if(Selection.Information(wdWithInTable)) {     MsgBox(Selection.Rows.SpaceBetweenColumns) } }` |

#### **Rows.TableDirection**

返回或设置 WPS 对指定的表格或行中的单元格进行排序的方向。可读写 **WdTableDirection** 类型。

**语法**

**express.TableDirection**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果将 **TableDirection** 属性设为 **wdTableDirectionLtr**，则选定的行按照最左侧的第一列进行排列。如果将 **TableDirection** 属性设为 **wdTableDirectionRtl**，则选定的行按照最右侧的第一列进行排列。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 从右向左对选定行中的单元格进行排序。*/ Selection.Rows.TableDirection = wdTableDirectionRtl` |

#### **Rows.VerticalPosition**

返回或设置由 **RelativeVerticalPosition** 属性指定的项目与行边缘之间的垂直距离。可读写 **Single** 类型。

**语法**

**express.VerticalPosition**

*express*   一个代表 **Rows** 对象的变量。

**说明**

该属性可以是一个以磅为单位表示度量的数字，也可以是任何有效的 **WdFramePosition** 常量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的第一个表格与页面顶端垂直对齐。*/ function test() { let myTable = ActiveDocument.Tables.Item(1).Rows myTable.RelativeVerticalPosition = wdRelativeVerticalPositionPage myTable.VerticalPosition = wdTableTop }` |

#### **Rows.WrapAroundText**

返回或设置文本是否环绕指定行。**Long** 类型，可读写。

**语法**

**express.WrapAroundText**

*express*   一个代表 **Rows** 对象的变量。

**说明**

如果设置仅环绕部分指定行，则返回 **wdUndefined**。可设置为 **True** 或 **False**。如果将 **WrapAroundText** 属性设置为 **False**，则 **AllowOverlap** 属性也会设置为 **False**。如果将 **AllowOverlap** 属性设置为 **True**，则 **WrapAroundText** 属性也会设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS ，使文本环绕在文档中第一张表格的周围。*/ ActiveDocument.Tables.Item(1).Rows.WrapAroundText = true` |

适用环境：web

适用平台：windows/linux