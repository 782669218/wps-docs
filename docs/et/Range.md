**Range**



代表某一单元格、某一行、某一列、某一选定区域（该区域可包含一个或若干连续单元格区域），或者某一三维区域。

**说明**

示例部分中说明了以下用于返回 **Range** 对象的属性和方法：

- **Range** 属性
- **Cells** 属性
- **Range** 和 **Cells**
- **Offset** 属性
- **Union** 方法

## 示例

使用 **Range**(*arg*)（其中 *arg* 为区域名称）可返回一个代表单个单元格或单元格区域的 **Range** 对象。下例将单元格 A1 中的值赋给单元格 A5。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Range("A5").Value2 = Application.Worksheets.Item("Sheet1").Range("A1").Value2 ` |

下例通过为区域 A1:H8 中的每个单元格设置公式，用随机数字填充该区域。如果在不带对象识别符（句点左边的对象）的情况下使用 **Range** 属性，该属性会返回活动表上的一个区域。如果活动表不是工作表，则该方法失败。在使用没有显式对象识别符的 **Range** 属性之前，请先使用 **Activate** 方法激活一个工作表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item("Sheet1").Activate()     //Range is on the active sheet     Application.Range("A1:H8").Formula = "=Rand()"  }` |

下例清除区域名为“*Criteria*”的区域中的内容。
注释：如果用文本参数指定区域地址，必须以 A1 样式记号指定该地址（不能用 R1C1 样式记号）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("Criteria").ClearContents()` |

使用 **Cells**(*row*, *column* )（其中 *row* 是行号，*column* 是列标）可返回一个单元格。下例将单元格 A1 赋值为 24。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Cells.Item(1, 1).Value2 = 24` |

下例设置单元格 A2 的公式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveSheet.Cells.Item(2, 1).Formula = "=Sum(B1:B5)"` |

虽然也可用 Range("A1") 返回单元格 A1，但有时用 **Cells** 属性更为方便，因为对行或列使用变量。下例在 Sheet1 上创建行号和列标。注意，当工作表激活以后，使用 **Cells** 属性时不必明确声明工作表（它将返回活动工作表上的单元格）。
注释：虽然可用 Visual Basic 字符串函数转换 A1 样式引用，但使用 Cells(1, 1) 记号更为简便（而且也是更好的编程习惯）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item("Sheet1").Activate()     for(let i = 1; i <= 5; i++)     {         Application.Cells.Item(1, i + 1).Value2 = 1990 + i      }       for(let j = 1; j <= 4; j++)     {         Application.Cells.Item(j + 1, 1).Value2 = "Q" + j     } }` |

使用 *expression*.**Cells**(*row* , *column*)（其中 *expression* 是返回 **Range** 对象的表达式，*row* 和 *column* 是相对于该区域左上角的偏移量）可返回区域中的一部分。下例设置单元格 C5 中的公式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("C5:C10").Cells.Item(1, 1).Formula = "=Rand()"` |

使用 **Range**(*cell1, cell2*)（其中 *cell1* 和 *cell2* 是指定起始和终止单元格的 **Range** 对象）可返回一个 **Range** 对象。下例设置单元格区域 A1:J10 的边框线条的样式。
注释：请注意每个 **Cells** 属性之前的句点。如果前导的 **With** 语句应用于 **Cells** 属性，那么这些句点就是必需的。本示例中，句点指示单元格处于工作表一上。如果没有句点，**Cells** 属性将返回活动工作表上的单元格。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range(Worksheets.Item(1).Cells.Item(1, 1),Application.Worksheets.Item(1).Cells.Item(10, 10)).Borders.LineStyle = xlThick` |

使用 **Offset**(*row, column*)（其中 *row* 和 *column* 为行偏移量和列偏移量）可返回相对于另一区域在指定偏移量处的区域。下例选定位于当前选定区域左上角单元格的向下三行且向右一列处的单元格。由于必须选定位于活动工作表上的单元格，因此必须先激活工作表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item("Sheet1").Activate()      //Can't select unless the sheet is active      Application.Selection.Offset(3, 1).Range("A1").Select() }` |

使用 **Union**(*range1, range2*, ...) 可返回多块区域，即该区域由两个或多个连续的单元格区域所组成。下例创建由单元格区域 A1:B2 和 C3:D4 组合定义的对象，然后选定该定义区域。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item("Sheet1").Activate()     let r1 = Range("A1:B2") let r2 = Range("C3:D4")     let myMultiAreaRange = Union(r1, r2)     Application.myMultiAreaRange.Select() }` |

如果您处理包含多个区域的选定内容，

[Areas](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#Range.Areas)

 属性是很有用的。它将多区域选定内容拆分为单个的 

Range

 对象，然后将对象返回为一个集合。您可以对返回的集合使用 

[Count](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#Range.Count)

 属性，以查找包含多个区域的选定内容，如下例所示。

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      let NumberOfSelectedAreas = Application.Selection.Areas.Count     if(NumberOfSelectedAreas > 1)     {         MsgBox("You cannot carry out this command " + "on multi-area selections")     } }` |

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**      | 激活单个单元格，该单元格必须处于当前选定区域内。要选择单元格区域，请使用 **Select** 方法。返回值为Variant类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Address**       | 返回一个 **String** 值，它代表宏语言的区域引用。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyNames**    | **表达式.ApplyNames(Names, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order, AppendLast)***表达式*   一个代表 **Range** 对象的变量。返回值为Variant类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFit**       | 更改区域中的列宽或行高以达到最佳匹配。返回值为Variant类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Calculate**     | 计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear**         | 清除整个对象。返回值为Variant类型。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearComments** | 清除指定区域的所有单元格批注。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearContents** | 清除区域中的公式。返回值类型为Variant。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormats**  | 清除对象的格式设置。返回值类型为Variant。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**          | 将单元格区域复制到指定的区域或剪贴板中。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateNames**   | 在指定区域中依据工作表中的文本标签创建名称。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CurrentRegion** | 返回一个 **Range** 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**           | 将对象剪切到剪贴板，或者将其粘贴到指定的目的地。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**        | 删除对象。返回值类型为Variant。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Find**          | 在区域中查找特定信息。返回值为一个 **Range** 对象，它代表第一个在其中找到该信息的单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindNext**      | 继续由 **Find** 方法开始的搜索。查找匹配相同条件的下一个单元格，并返回表示该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。返回值类型为Range。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindPrevious**  | 继续由 **Find** 方法开始的搜索。查找匹配相同条件的上一个单元格，并返回代表该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。返回值类型为Range。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Insert**        | 在工作表或宏表中插入一个单元格或单元格区域，其他单元格相应移位以腾出空间。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**          | 返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表对指定区域某一偏移量处的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Justify**       | 调整区域内的文字，使之均衡地填充该区域。返回值类型为Variant。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge**         | 由指定的 **Range** 对象创建合并单元格。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Offset**        | 返回 **Range** 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintPreview**  | 按对象打印后的外观效果显示对象的预览。返回值类型为Variant。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**         | 返回一个 **Range** 对象，它代表一个单元格或单元格区域。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Resize**        | 调整指定区域的大小。返回 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表调整后的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Run**           | 在该处运行 ET 宏。区域必须位于宏表上。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**        | 选择对象。返回值类型为Variant。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UnMerge**       | 将合并区域分解为独立的单元格。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Value**         | 返回一个 **Variant** 型，它代表指定单元格的值。              |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddIndent**           | 返回或设置一个 **Variant** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Areas**               | 返回一个 **Areas** 集合，该集合表示多重区域选择中的所有区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Cells**               | 返回一个 Range 对象，该对象代表指定区域中的单元格。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Column**              | 返回指定区域中第一块中的第一列的列号。Long 类型，只读。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnWidth**         | 返回或设置指定区域中所有列的列宽。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Columns**             | 返回一个 **Range** 对象，该对象代表指定区域中的列。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**               | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Dependents**          | 返回一个 **Range** 对象，该对象表示包含一个单元格所有从属单元格的区域。如果有多个从属单元格，这可能有多个选择（多个 **Range** 对象）。**Range** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **End**                 | 返回一个 **Range** 对象，该对象代表包含源区域的区域尾端的单元格。等同于按键 End+ 向上键、End+ 向下键、End+ 向左键或 End+ 向右键。**Range** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EntireColumn**        | 返回一个 **Range** 对象，该对象表示包含指定区域的整列（或多列）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EntireRow**           | 返回一个 **Range** 对象，该对象表示包含指定区域的整行（或多行）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**                | 返回一个 **Font**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象，它代表指定对象的字体。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**             | 返回或设置一个 **Variant** 值，它代表 A1 样式表示法和宏语言中的对象的公式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 返回或设置一个 **Variant** 值，该值代表区域的高度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**              | 返回或设置一个 **Variant** 值，它指明是否隐藏行或列。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAlignment** | 返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                  | 返回或设置一个 **String** 值，它代表将页面另存为网页时指定单元格的识别标志。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                | 返回一个 **Variant** 值，它代表从列 A 的左边缘到区域的左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MergeArea**           | 返回一个 **Range** 对象，该对象代表包含指定单元格的合并区域。如果指定的单元格不在合并区域内，则该属性返回指定的单元格。只读。**Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MergeCells**          | 如果区域包含合并单元格，则为 **True**。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回或设置一个 **Variant** 值，它代表对象的名称。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat**        | 返回或设置一个 **Variant** 值，它代表对象的格式代码。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**            | 返回一个代表下一个单元格的 **Range** 对象。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Row**                 | 返回区域中第一个子区域的第一行的行号。**Long** 类型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowHeight**           | 以磅为单位返回或设置指定区域中所有行的行高。如果指定区域中的各行的行高不等，则返回 **null**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rows**                | 返回一个 **Range** 对象，它代表指定单元格区域中的行。**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**                | 返回或设置指定对象中的文本。**String** 型，只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                 | 返回或设置一个 **Variant** 值，它代表行 1 上边缘到区域上边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value2**              | 返回或设置单元格值。**Variant** 类型，可读写。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 返回一个 **Variant** 值，它代表区域的宽度（以单位表示）。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Worksheet**           | 返回一个 **Worksheet** 对象，该对象表示包含指定区域的工作表。只读。 |

**成员方法**

#### **Range.Activate**

激活单个单元格，该单元格必须处于当前选定区域内。要选择单元格区域，请使用 **Select** 方法。返回值为Variant类型。

**语法**

**express.Activate()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定工作表 Sheet1 上的单元格区域 A1:C3，并激活单元格 B2。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Range("A1:C3").Select()     Application.Range("B2").Activate() }` |

#### **Range.Address**

返回一个 **String** 值，它代表宏语言的区域引用。

**语法**

**express.Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External, RelativeTo)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**         | **说明**                                                     |
| ---------------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *RowAbsolute*    | 可选          | **Variant**          | 如果为 True，则以绝对引用返回引用的行部分。默认值为 True。   |
| *ColumnAbsolute* | 可选          | **Variant**          | 如果为 True，则以绝对引用返回引用的列部分。默认值为 True。   |
| *ReferenceStyle* | 可选          | **XlReferenceStyle** | 引用样式。默认值为 xlA1。                                    |
| *External*       | 可选          | **Variant**          | 如果为 True，则返回外部引用。如果为 False，则返回本地引用。默认值为 False。 |
| *RelativeTo*     | 可选          | **Variant**          | 如果 RowAbsolute 和 ColumnAbsolute 为 False，并且 ReferenceStyle 为 xlR1C1，则必须包括相对引用的起始点。此参数是定义起始点的 Range 对象。 |

**说明**

如果引用包含多个单元格，*RowAbsolute* 和 *ColumnAbsolute* 将应用于所有的行和列。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例对工作表 Sheet1 中的同一单元格地址使用了四种不同的表达方式。示例中的注释为将要显示于消息框中的地址。*/ function test() {     let mc = Application.Worksheets.Item("Sheet1").Cells.Item(1, 1)     MsgBox(mc.Address())                                    // $A$1     MsgBox(mc.Address(false))                               // $A1     MsgBox(mc.Address(undefined,undefined,xlR1C1))          // R1C1     MsgBox(mc.Address(false,false,xlR1C1,undefined,Worksheets.Item(1).Cells.Item(3, 3)))      // R[-2]C[-2] }` |

#### **Range.ApplyNames**

**表达式.ApplyNames(Names, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order, AppendLast)**

*表达式*   一个代表 **Range** 对象的变量。

返回值为Variant类型。

**语法**

**express.ApplyNames(Names, IgnoreRelativeAbsolute, UseRowColumnNames, OmitColumn, OmitRow, Order, AppendLast)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**          | **说明**                                                     |
| ------------------------ | ------------- | --------------------- | ------------------------------------------------------------ |
| *Names*                  | 可选          | **Variant**           | 要应用的名称数组。如果省略该参数，则工作表中所有的名称都将应用到该区域上。 |
| *IgnoreRelativeAbsolute* | 可选          | **Variant**           | 要应用的名称数组。如果省略该参数，则工作表中所有的名称都将应用到该区域上。 如果为 True，则用名称取代引用，而不管名称或引用的类型如何。如果为 False，则只用绝对名称替换绝对引用，用相对名称替换相对引用，并只用混合名称替换混合引用。默认值为 True。 |
| *UseRowColumnNames*      | 可选          | **Variant**           | 如果为 True，则当无法找到指定区域的名称时，使用该区域所在行或列区域的名称。如果为 False，则忽略 OmitColumn 和 OmitRow 参数。默认值为 True。 |
| *OmitColumn*             | 可选          | **Variant**           | 如果为 True，则用行方向的名称替换整个引用。仅当引用的单元格与公式位于同一列中，且处于行方向命名的区域中时，才能省略列方向名称。默认值为 True。 |
| *OmitRow*                | 可选          | **Variant**           | 如果为 True，则用列方向的名称替换整个引用。仅当引用的单元格与公式位于同一行中，且处于列方向命名的区域中时，才能省略行方向名称。默认值为 True。 |
| *Order*                  | 可选          | **XlApplyNamesOrder** | 确定用行方向区域名称和列方向区域名称替换单元格引用时，首先列出哪个区域名称。 |
| *AppendLast*             | 可选          | **Variant**           | 如果为 True，则替换 Names 中的名称定义以及所定义的姓氏的定义。如果为 False，则只替换 Names 中的名称定义。默认值为 False。 |

**说明**

可用 **Array** 函数为 *Names* 参数创建名称列表。

如果要对整个工作表应用名称，可用 Cells.ApplyNames。

不能“取消应用”名称；要删除名称，请使用 **Delete** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对整个工作表应用名称。*/ Application.Cells.ApplyNames(["Sales","Profits"])` |

#### **Range.AutoFit**

更改区域中的列宽或行高以达到最佳匹配。

返回值为Variant类型。

**语法**

**express.AutoFit()**

*express*   一个代表 **Range** 对象的变量。

**说明**

**Range** 对象必须是行或行区域，或者列或列区域。否则，该方法将产生错误。

一个列宽单位等于“常规”样式中一个字符的宽度。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例调整工作表 Sheet1 中从 A 到 I 的列，以获得最适当的列宽。*/ function test() {     Application.Worksheets.Item("Sheet1").Columns.Item("A:I").AutoFit() }  /*本示例调整工作表 Sheet1 上从 A 到 E 的列，以获得最适当的列宽，但该调整仅依据单元格区域 A1:E1 中的内容进行。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1:E1").Columns.AutoFit() }` |

#### **Range.Calculate**

计算所有打开的工作簿、工作簿中的某张特定工作表或工作表指定区域中的单元格，如下表所示。

返回值类型为Variant。

**语法**

**express.Calculate()**

*express*   一个代表 **Range** 对象的变量。

**说明**

| 要计算           | 依照此示例                                   |
| ---------------- | -------------------------------------------- |
| 所有打开的工作簿 | Application.Calculate()（或只是Calculate()） |
| 指定工作表       | Worksheets.Item(1).Calculate()               |
| 指定区域         | Worksheets.Item(1).Rows.Item(2).Calculate()  |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例计算 Sheet1 上已用区域中 A 列、B 列和 C 列的公式。*/ function test() {     Application.Worksheets.Item("Sheet1").UsedRange.Columns.Item("A:C").Calculate() }` |

#### **Range.Clear**

清除整个对象。

返回值为Variant类型。

**语法**

**express.Clear()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例清除工作表 Sheet1 中 A1:G37 单元格区域的公式和格式设置。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1:G37").Clear() }` |

#### **Range.ClearComments**

清除指定区域的所有单元格批注。

**语法**

**express.ClearComments()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除 E5 单元格的所有批注。*/ Application.Worksheets.Item(1).Range("E5").ClearComments()` |

#### **Range.ClearContents**

清除区域中的公式。

返回值类型为Variant。

**语法**

**express.ClearContents()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的公式，但保留其格式设置。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearContents() }` |

#### **Range.ClearFormats**

清除对象的格式设置。

返回值类型为Variant。

**语法**

**express.ClearFormats()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例清除工作表 Sheet1 上 A1:G37 单元格区域的所有格式设置。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1:G37").ClearFormats() }  /*此示例清除工作表 Sheet1 上第一个嵌入式图表的格式设置。*/ function test() {     Application.Worksheets.Item("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats() }` |

#### **Range.Copy**

将单元格区域复制到指定的区域或剪贴板中。

返回值类型为Variant。

**语法**

**express.Copy(Destination)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Destination* | 可选          | **Variant**  | 指定区域要复制到的新域。如果省略此参数，ET 会将区域复制到剪贴板。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将工作表 Sheet1 上单元格区域 A1:D4 中的公式复制到工作表 Sheet2 上的单元格区域 E5:H8 中。*/ Application.Worksheets.Item("Sheet1").Range("A1:D4").Copy(Worksheets.Item("Sheet2").Range("E5")) ` |

#### **Range.CreateNames**

在指定区域中依据工作表中的文本标签创建名称。

返回值类型为Variant。

**语法**

**express.CreateNames(Top, Left, Bottom, Right)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Top*    | 可选          | **Variant**  | 如果该值为 True，则使用顶部行中的标签创建名称。默认值为 False。 |
| *Left*   | 可选          | **Variant**  | 如果该值为 True，则使用左侧列中的标签创建名称。默认值为 False。 |
| *Bottom* | 可选          | **Variant**  | 如果该值为 True，则使用底部行中的标签创建名称。默认值为 False。 |
| *Right*  | 可选          | **Variant**  | 如果该值为 True，则使用右侧列中的标签创建名称。默认值为 False。 |

**说明**

如果未指定 *Top*、*Left*、*Bottom* 或 *Right* 之一，ET 将根据指定区域的形状猜测文本标志。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用单元格区域 A1:A3 中的文字创建区域 B1:B3 的名称。注意，指定区域时必须包括那些含有名称文字的单元格，即便只是为单元格区域 B1:B3 创建名称。*/ function test() {     let rangeToName = Application.Worksheets.Item("Sheet1").Range("A1:B3")     rangeToName.CreateNames(null, true) }` |

#### **Range.CurrentRegion**

返回一个 **Range** 对象，该对象表示当前区域。当前区域是以空行与空列的组合为边界的区域。只读。

**语法**

**express.CurrentRegion()**

*express*   一个代表 **Range** 对象的变量。

**说明**

该属性对于许多自动展开选择以包括整个当前区域的操作很有用，如 **AutoFormat** 方法。

该属性不能用于被保护的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定工作表 Sheet1 上的当前区域。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.CurrentRegion.Select() }  /*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/ function test() {     let tbl = Application.ActiveCell.CurrentRegion     tbl.Offset(1, 0).Resize(tbl.Rows.Count-1,tbl.Columns.Count).Select() }` |

#### **Range.Cut**

将对象剪切到剪贴板，或者将其粘贴到指定的目的地。

返回值类型为Variant。

**语法**

**express.Cut(Destination)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Destination* | 可选          | **Variant**  | 应在其中粘贴对象的目标区域。如果省略此参数，区域对象会被剪切到剪贴板。 |

**说明**

剪切的区域必须由相邻的单元格组成。

#### **Range.Delete**

删除对象。

返回值类型为Variant。

**语法**

**express.Delete(Shift)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Shift*  | 可选          | **Variant**  | 仅用于 Range 对象。指定如何调整单元格以替换删除的单元格。可以为以下 XlDeleteShiftDirection 常量之一：xlShiftToLeft 或 xlShiftUp。如果省略此参数，ET 将根据区域的形状确定调整方式。 |

#### **Range.Find**

在区域中查找特定信息。

返回值为一个 **Range** 对象，它代表第一个在其中找到该信息的单元格。

**语法**

**express.Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase, MatchByte, SearchFormat)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**          | **说明**                                                     |
| ----------------- | ------------- | --------------------- | ------------------------------------------------------------ |
| *What*            | 可选          | **Variant**           | 要搜索的数据。可为字符串或任意 ET 数据类型。                 |
| *After*           | 可选          | **Variant**           | 表示搜索过程将从其之后开始进行的单元格。此单元格对应于从用户界面搜索时的活动单元格的位置。请注意：After 必须是区域中的单个单元格。要记住搜索是从该单元格之后开始的；直到此方法绕回到此单元格时，才对其进行搜索。如果不指定该参数，搜索将从区域的左上角的单元格之后开始。 |
| *LookIn*          | 可选          | **Variant**           | 信息类型。                                                   |
| *LookAt*          | 可选          | **Variant**           | 可为以下 XlLookAt 常量之一：xlWhole 或 xlPart。              |
| *SearchOrder*     | 可选          | **Variant**           | 可为以下 XlSearchOrder 常量之一：xlByRows 或 xlByColumns。   |
| *SearchDirection* | 可选          | **XlSearchDirection** | 搜索的方向。                                                 |
| *MatchCase*       | 可选          | **Variant**           | 如果为 True，则搜索区分大小写。默认值为 False。              |
| *MatchByte*       | 可选          | **Variant**           | 只在已经选择或安装了双字节语言支持时适用。如果为 True，则双字节字符只与双字节字符匹配。如果为 False，则双字节字符可与其对等的单字节字符匹配。 |
| *SearchFormat*    | 可选          | **Variant**           | 搜索的格式。                                                 |

**说明**

如果未发现匹配项，则返回 **Nothing**。**Find** 方法不影响选定区域或当前活动的单元格。

每次使用此方法后，参数 *LookIn*、*LookAt*、*SearchOrder* 和 *MatchByte* 的设置都将被保存。如果下次调用此方法时不指定这些参数的值，就使用保存的值。设置这些参数将更改**“查找”**对话框中的设置，如果省略这些参数，更改**“查找”**对话框中的设置将更改使用的保存值。要避免出现这一问题，每次使用此方法时请明确设置这些参数。

使用 **FindNext** 和 **FindPrevious** 方法可重复搜索。

当搜索到指定查找区域的末尾时，此方法将绕回到区域的开始位置继续搜索。发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。

若要对单元格进行模式更为复杂的搜索，请结合使用 For Each...Next 语句和 **Like** 运算符。例如，下列代码在单元格区域 A1:C5 中搜索字体名称以“Cour”开始的单元格。当 ET 找到匹配单元格以后，就将其字体改为 Times New Roman。

For Each c In [A1:C5] If c.Font.Name Like "Cour*" Then c.Font.Name = "Times New Roman" End If Next

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在第一个工作表的单元格区域 A1:A500 中查找包含值 2 的所有单元格，并将这些单元格的值更改为 5。*/ function test() {     let add = Application.Worksheets.Item(1).Range("A1:A500")      let c = add.Find(2,undefined,xlValues)     if(c != null) {         let firstAddress = c.Address         do {             c.Value2 = 5             c = add.FindNext(c)         }while(c != null && c.Address != firstAddress)     } }` |

#### **Range.FindNext**

继续由 **Find** 方法开始的搜索。查找匹配相同条件的下一个单元格，并返回表示该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。

返回值类型为Range。

**语法**

**express.FindNext(After)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *After*  | 可选          | **Variant**  | 指定一个单元格，查找将从该单元格之后开始。此单元格对应于从用户界面搜索时的活动单元格位置。注意，After 必须是查找区域中的单个单元格。注意，搜索是从该单元格之后开始的；直到本方法环绕到此单元格时，才检测其内容。如果未指定本参数，查找将从区域的左上角单元格之后开始。 |

**说明**

当查找到指定查找区域的末尾时，本方法将环绕至区域的开始继续搜索。发生环绕后，为停止查找，可保存第一次找到的单元格地址，然后测试下一个查找到的单元格地址是否与其相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在单元格区域 A1:A500 中查找值为 2 的单元格，并将这些单元格的值变为 5。*/ function test() {     let add = Application.Worksheets.Item(1).Range("A1:A500")      let c = add.Find(2,undefined,xlValues)     if(c != null) {         let firstAddress = c.Address         do {     	    c.Value2 = 5     	    c = add.FindNext(c)         }while(c != null && c.Address != firstAddress)     } }` |

#### **Range.FindPrevious**

继续由 **Find** 方法开始的搜索。查找匹配相同条件的上一个单元格，并返回代表该单元格的 **Range** 对象。该操作不影响选定内容和活动单元格。

返回值类型为Range。

**语法**

**express.FindPrevious(After)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *After*  | 可选          | **Variant**  | 指定一个单元格，查找将从该单元格之前开始。此单元格对应于从用户界面搜索时的活动单元格的位置。请注意：After 必须是区域中的单个单元格。注意，搜索是从该单元格之前开始的；直到本方法环绕到此单元格时，才检测其内容。如果未指定本参数，查找将从区域的左上角单元格之前开始。 |

**说明**

当查找到指定查找区域的起始位置时，本方法将环绕至区域的末尾继续搜索。发生绕回后，要停止搜索，可保存第一个找到的单元格地址，然后测试后面找到的每个单元格地址是否与其相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例演示 FindPrevious 方法如何与 Find 方法和 FindNext 方法一起使用。运行本示例之前，请确保工作表 Sheet1 的 B 列中至少出现过两次“Phoenix”单词。*/ function test() {     let fc = Application.Worksheets.Item("Sheet1").Columns.Item("B").Find("Phoenix")         MsgBox("The first occurrence is in cell " + fc.Address)     let Mc = Application.Worksheets.Item("Sheet1").Columns.Item("B").FindNext(fc)         MsgBox("The next occurrence is in cell " + Mc.Address)     let gc = Application.Worksheets.Item("Sheet1").Columns.Item("B").FindPrevious(fc)         MsgBox("The previous occurrence is in cell " + gc.Address) }` |

#### **Range.Insert**

在工作表或宏表中插入一个单元格或单元格区域，其他单元格相应移位以腾出空间。

返回值类型为Variant。

**语法**

**express.Insert(Shift, CopyOrigin)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Shift*      | 可选          | **Variant**  | 指定单元格的调整方式。可为以下 XlInsertShiftDirection 常量之一：xlShiftToRight 或 xlShiftDown。如果省略此参数，ET 将根据区域的形状确定调整方式。 |
| *CopyOrigin* | 可选          | **Variant**  | 复制的起点。                                                 |

#### **Range.Item**

返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表对指定区域某一偏移量处的区域。

**语法**

**express.Item(RowIndex, ColumnIndex)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *RowIndex*    | 必选          | **Variant**  | 要访问的单元格的索引号，顺序为从左到右，然后往下。 Range.Item(1) 返回区域左上角单元格； Range.Item(2) 返回紧靠左上角单元格右侧的单元格。 |
| *ColumnIndex* | 可选          | **Variant**  | 指明要访问的单元格所在列的列号的数字或字符串，1 或 “A”表示区域中的第一列。 |

**说明**

语法 1 使用行号和列号或列标作为索引参数。关于此语法的详细信息，请参阅 **Range** 对象。**RowIndex** 和 **ColumnIndex** 参数为相对偏移量。换句话说，如果 **RowIndex** 指定为 1，则返回区域内第一行中的单元格，而非工作表的第一行。例如，如果选定区域为单元格 C3，则 Selection.Cells(2, 2) 返回单元格 D4（使用 **Item** 属性可在原始区域之外进行索引）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例基于单元格 A1 的内容填写 Sheet1 的单元格区域 A2。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1").Item(2).FillDown()	 }` |

#### **Range.Justify**

调整区域内的文字，使之均衡地填充该区域。

返回值类型为Variant。

**语法**

**express.Justify()**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果区域不够大，ET 将显示一个消息，通知您文本将延伸到区域的下方。如果单击“确定”按钮，被调整的文本将替换超出选定区域的单元格的内容。为防止该消息出现，请将 **DisplayAlerts** 属性值设置为 **False**。设置该属性后，文本将始终替换区域下方单元格的内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例调整工作表 Sheet1 上单元格 A1 中的文字。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1").Justify() }` |

#### **Range.Merge**

由指定的 **Range** 对象创建合并单元格。

**语法**

**express.Merge(Across)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Across* | 可选          | **Variant**  | 如果为 True，则将指定区域中每一行的单元格合并为一个单独的合并单元格。默认值是 False。 |

**说明**

合并区域的值在该区域左上角的单元格中指定。

#### **Range.Offset**

返回 **Range** 对象，它代表位于指定单元格区域的一定的偏移量位置上的区域。

**语法**

**express.Offset(RowOffset, ColumnOffset)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *RowOffset*    | 可选          | **Variant**  | 区域偏移的行数（正数、负数或 0（零））。正数表示向下偏移，负数表示向上偏移。默认值是 0。 |
| *ColumnOffset* | 可选          | **Variant**  | 区域偏移的列数（正数、负数或 0（零））。正数表示向右偏移，负数表示向左偏移。默认值是 0。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例激活 Sheet1 上活动单元格向右偏移三列、向下偏移三行处的单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.Offset(3,3).Activate() }  /*此示例假设 Sheet1 中包含一个具有标题行的表格。此示例先选定该表格，但并不选择行首。运行此示例之前，活动单元格必须位于表格中。*/ function test() {     let tbl = Application.ActiveCell.CurrentRegion     tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1,tbl.Columns.Count).Select() }` |

#### **Range.PrintPreview**

按对象打印后的外观效果显示对象的预览。

返回值类型为Variant。

**语法**

**express.PrintPreview(EnableChanges)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EnableChanges* | 可选          | **Variant**  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在打印预览中显示工作表 Sheet1。*/ Application.Worksheets.Item("Sheet1").PrintPreview()` |

#### **Range.Range**

返回一个 **Range** 对象，它代表一个单元格或单元格区域。

**语法**

**express.Range(Cell1, Cell2)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cell1*  | 必选          | **Variant**  | 区域名称。必须为采用宏语言的 A1 样式引用。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可包括货币符号，但它们会被忽略掉。您可以在区域中任一部分使用局部定义名称。如果使用名称，则假定该名称使用的是宏语言。 |
| *Cell2*  | 可选          | **Variant**  | 区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，或者也可以是一个用宏语言为单个单元格命名的字符串。 |

**说明**

如果在没有对象识别符时使用，则该属性是 ActiveSheet.Range 的快捷方式（它返回活动表的一个区域，如果活动表不是一张工作表，则该属性无效）。

当应用于 **Range** 对象时，该属性与 **Range** 对象相关。例如，如果选中单元格 C3，那么 Selection.Range("B1") 返回单元格 D3，因为它同 **Selection** 属性返回的 **Range** 对象相关。此外，代码 ActiveSheet.Range("B1") 总是返回单元格 B1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将 Sheet1 上 A1 单元格的值设置为 3.14159。*/ Application.Worksheets.Item("Sheet1").Range("A1").Value2 = 3.14159  /*此示例在 Sheet1 的 A1 单元格中创建一个公式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=10*RAND()"  /*此示例在 Sheet1 上的单元格区域 A1:D10 中进行循环。如果某个单元格的值小于 0.001，则此代码将用 0（零）来取代该值。*/ function test() {     let c = Application.Worksheets.Item("Sheet1").Range("A1:D10")     for(let i = 1; i <= c.Rows.Count;i++) {         for(let j = 1 ;j <= c.Columns.Count;j++) {             if(c.Item(i,j).Value2 < 0.001) {                 c.Item(i,j).Value2 = 0             }         }     } }  /*此示例在名为“TestRange”的区域上进行循环，并显示该区域中空白单元格的个数。*/ function test() {         let numBlanks = 0     let c = Application.Range("TestRange")     for(let i = 1 ; i <= c.Rows.Count; i++) {         for(let j = 1; j <= c.Columns.Count; j++) {             if(c.Item(i,j).Value2 == null) {                 numBlanks = numBlanks + 1             }         }     }     MsgBox("There are " + numBlanks + " empty cells in this range") }  /*此示例将 Sheet1 中单元格区域 A1:C5 上的字体样式设置为斜体。此示例使用 Range 属性的语法 2。*/ Application.Worksheets.Item("Sheet1").Range(Cells.Item(1, 1), Cells.Item(5, 3)).Font.Italic = true` |

#### **Range.Resize**

调整指定区域的大小。返回 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表调整后的区域。

**语法**

**express.Resize(RowSize, ColumnSize)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                   |
| ------------ | ------------- | ------------ | ---------------------------------------------------------- |
| *RowSize*    | 可选          | **Variant**  | 新区域中的行数。如果省略该参数，则该区域中的行数保持不变。 |
| *ColumnSize* | 可选          | **Variant**  | 新区域中的列数。如果省略该参数。则该区域中的列数保持不变。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例调整 Sheet1 中选定区域的大小，使之增加一行和一列。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     numRows = Application.Selection.Rows.Count     numColumns = Application.Selection.Columns.Count     Application.Selection.Resize(numRows + 1, numColumns + 1).Select() }  /*本示例假定在 Sheet1 中有一个包含标题行的表。本示例选定该表，但不选定标题行。运行本示例之前，活动单元格必须处于该表中。*/ function test() {       let tbl = Application.ActiveCell.CurrentRegion     tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1,tbl.Columns.Count).Select() }` |

#### **Range.Run**

在该处运行 ET 宏。区域必须位于宏表上。

**语法**

**express.Run(Arg1, ..., Arg30)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Arg1*   | 可选          | **Variant**  | 传递给函数的参数。 |
| *...*    | 可选          | **Variant**  | 传递给函数的参数。 |
| *Arg30*  | 可选          | **Variant**  | 传递给函数的参数。 |

**说明**

此方法不可使用命名参数，参数必须通过位置进行传递。

**Run** 方法返回被调用的宏返回的任何值。如果将对象作为参数传递给宏，该对象将转换为相应的值（通过对该对象应用 **Value** 属性）。这意味着不能用 **Run** 方法将对象传递给宏。

#### **Range.Select**

选择对象。

返回值类型为Variant。

**语法**

**express.Select()**

*express*   一个代表 **Range** 对象的变量。

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用 **Activate** 方法。

#### **Range.UnMerge**

将合并区域分解为独立的单元格。

**语法**

**express.UnMerge()**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将包含单元格 A3 的合并区域分解。*/ function test() {     if(Application.Range("a3").MergeCells){         Application.Range("a3").MergeArea.UnMerge()     }     else {         MsgBox( "not merged")     } }` |

#### **Range.Value**

返回一个 **Variant** 型，它代表指定单元格的值。

**语法**

**express.Value(RangeValueDataType)**

*express*   一个代表 **Range** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                           |
| -------------------- | ------------- | ------------ | -------------------------------------------------- |
| *RangeValueDataType* | 可选          | **Variant**  | 区域值数据类型。可以为 xlRangeValueDataType 常量。 |

**说明**

用 XML 电子表格文件中的内容设置单元格区域时，只使用工作簿中第一张工作表中的值。无法设置或得到以 XML 电子表格格式表示的不相连的单元格区域。

JSAPI里的Value是个方法，可以进行取值，不能进行赋值。要给单元格赋值，请用Range.Value2属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例获取 Sheet1 上 A1 单元格的值，赋给变量"rngValue"。*/ let rngValue = Application.Worksheets.Item("Sheet1").Range("A1").Value()` |

**成员属性**

#### **Range.AddIndent**

返回或设置一个 **Variant** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果将此属性的值设为 **True**，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 **Orientation** 属性的值为 **xlVertical** 时，将 **VerticalAlignment** 属性设为 **xlVAlignDistributed**；在 **Orientation** 属性的值为 **xlHorizontal** 时，将 **HorizontalAlignment** 属性设为 **xlHAlignDistributed**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将 Sheet1 的单元格 A1 中文本的水平对齐方式设置为等距分布，然后缩进该文本。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A1").HorizontalAlignment = xlHAlignDistributed     Application.Worksheets.Item("Sheet1").Range("A1").AddIndent = true }` |

#### **Range.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = ActiveWorkbook     if (myObject.Application.Value == "ET") {         MsgBox("This is an ET Application object.")     }     else {         MsgBox("This is not an ET Application object.")     } }` |

#### **Range.Areas**

返回一个 **Areas** 集合，该集合表示多重区域选择中的所有区域。只读。

**语法**

**express.Areas**

*express*   一个代表 **Range** 对象的变量。

**说明**

对于单一选择区域，**Areas** 属性返回只包含一个对象的集合，即原始 **Range** 对象本身。对于多重选择区域，**Areas** 属性返回一个集合，该集合包含与每个选定区域相对应的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在用户选定多个区域并试图执行某一命令时显示提示信息。该示例必须在工作表上运行。*/ function test() {     if(Application.Selection.Areas.Count > 1) {         MsgBox("Cannot do this to a multi-area selection.")     } }` |

#### **Range.Cells**

返回一个 Range 对象，该对象代表指定区域中的单元格。

**语法**

**express.Cells**

*express*   一个代表 **Range** 对象的变量。

**说明**

因为 **Item** 属性是 **Range** 对象的默认属性，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 **Range** 对象，它代表活动工作表中所有的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 上单元格区域 A1: C5 的字体样式设置为斜体。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Worksheets.Item("Sheet1").Range(Cells.Item(1, 1), Cells.Item(5, 3)).Font.Italic = true }  /*本示例搜索列“myRange”中的数据。如果发现某单元格的值与上面的一个单元格的值相等，则本示例将显示这个包含重复数据的单元格的地址。*/ function test() {     let r = Range("myRange")     for(let n = 1;n<=r.Rows.Count;n++) {         if(r.Cells.Item(n, 1).Value2 == r.Cells.Item(n + 1, 1).Value2) {             MsgBox("Duplicate data in " + r.Cells.Item(n + 1, 1).Address())         }     } }` |

#### **Range.Column**

返回指定区域中第一块中的第一列的列号。Long 类型，只读。

**语法**

**express.Column**

*express*   一个代表 **Range** 对象的变量。

**说明**

A 列返回 1，B 列返回 2，依此类推。

若要返回区域中最后一列的列号，请使用下列语句。

myRange.Columns(myRange.Columns.Count).Column

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作表 Sheet1 上每隔一列的列宽设置为 4 磅。*/ function test() {     let col = Application.Worksheets.Item("Sheet1").Columns     for (let i = 1; i <= col.Count; i++) {         if ((col.Item(i).Column) % 2 == 0) {             col.Item(i).ColumnWidth = 4         }     } }` |

#### **Range.ColumnWidth**

返回或设置指定区域中所有列的列宽。**Variant** 类型，可读写。

**语法**

**express.ColumnWidth**

*express*   一个代表 **Range** 对象的变量。

**说明**

一个列宽单位等于“常规”样式中一个字符的宽度。对于比例字体，则使用字符 0（零）的宽度。

使用 **Width** 属性可返回以磅为单位的列宽。

如果区域中所有列的列宽都相等，**ColumnWidth** 属性返回该宽度值。如果区域中的列宽不等，该属性返回 **null**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使工作表 Sheet1 上 A 列的列宽加倍。*/ function test() {     let itm = Application.Worksheets.Item("Sheet1").Columns.Item("A")     itm.ColumnWidth = itm.ColumnWidth * 2 }` |

#### **Range.Columns**

返回一个 **Range** 对象，该对象代表指定区域中的列。

**语法**

**express.Columns**

*express*   一个代表 **Range** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ActiveSheet.Columns。

此属性在应用于一个是多重选定区域的 **Range** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **Range** 对象有两个子区域 A1:B2 和 C3:D4，那么，Selection.Columns.Count 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 Areas.Count 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将名为“myRange”区域第一列中每一单元格的值置为 0（零〕。*/ Application.Range("myRange").Columns.Item(1).Value2 = 0  /*此示例显示 Sheet1 上选定区域中的列数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let areaCount = Application.Selection.Areas.Count     MsgBox(areaCount)     if (areaCount <= 1) {         MsgBox("The selection contains " + Application.Selection.Columns.Count + " columns.")     }     else {         for (let i = 1; i <= areaCount; i++) {             MsgBox("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")         }     } }` |

#### **Range.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Sheet1 上选定区域中的列数。此示例还将检测选定区域中是否包含多重选定区域，如果包含，则对多重选定区域中每一子区域进行循环。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let iAreaCount = Application.Selection.Areas.Count     if(iAreaCount <= 1) {         MsgBox("The selection contains " + Application.Selection.Columns.Count + " columns.")     }     else {         for(let i = 1;i<=iAreaCount;i++) {             MsgBox("Area " + i + " of the selection contains " + Application.Selection.Areas.Item(i).Columns.Count + " columns.")         }     } }` |

#### **Range.Dependents**

返回一个 **Range** 对象，该对象表示包含一个单元格所有从属单元格的区域。如果有多个从属单元格，这可能有多个选择（多个 **Range** 对象）。**Range** 类型，只读。

**语法**

**express.Dependents**

*express*   一个代表 **Range** 对象的变量。

**说明**

**Dependents** 属性仅在活动工作表中有效，不能追踪远程引用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定工作表 Sheet1 中单元格 A1 的从属单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Range("A1").Dependents.Select() }` |

#### **Range.End**

返回一个 **Range** 对象，该对象代表包含源区域的区域尾端的单元格。等同于按键 End+ 向上键、End+ 向下键、End+ 向左键或 End+ 向右键。**Range** 对象，只读。

**语法**

**express.End**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定包含单元格 B4 的区域中 B 列顶端的单元格。*/ Application.Range("B4").End(xlUp).Select()  /*本示例选定包含单元格 B4 的区域中第 4 行尾端的单元格。*/ Application.Range("B4").End(xlToRight).Select()  /*本示例将选定区域从单元格 B4 延伸至第四行最后一个包含数据的单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.Range("B4", Range("B4").End(xlToRight)).Select() }` |

#### **Range.EntireColumn**

返回一个 **Range** 对象，该对象表示包含指定区域的整列（或多列）。只读。

**语法**

**express.EntireColumn**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对包含活动单元格的列中的第一个单元格赋值。本示例必须在工作表上运行。*/ Application.ActiveCell.EntireColumn.Cells.Item(1, 1).Value2 = 5` |

#### **Range.EntireRow**

返回一个 **Range** 对象，该对象表示包含指定区域的整行（或多行）。只读。

**语法**

**express.EntireRow**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对包含活动单元格的行中的第一个单元格赋值。本示例必须在工作表上运行。*/ Application.ActiveCell.EntireRow.Cells.Item(1, 1).Value2 = 5` |

#### **Range.Font**

返回一个 **Font**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象，它代表指定对象的字体。

**语法**

**express.Font**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例判断单元格 A1 的字体名称是否为 Arial，并通知用户。*/ function CheckFont(){     Application.Range("A1").Select() 								     // Determine if the font name for selected cell is Arial.     if(Application.Range("A1").Font.Name == "Arial") {         MsgBox("The font name for this cell is 'Arial'")     }     else {         MsgBox("The font name for this cell is not 'Arial'")     } }` |

#### **Range.Formula**

返回或设置一个 **Variant** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

*express*   一个代表 **Range** 对象的变量。

**说明**

此属性对于 OLAP 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Sheet1 中 A1 单元格的公式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=$A$4+$A$10"` |

#### **Range.Height**

返回或设置一个 **Variant** 值，该值代表区域的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **Range** 对象的变量。

#### **Range.Hidden**

返回或设置一个 **Variant** 值，它指明是否隐藏行或列。

**语法**

**express.Hidden**

*express*   一个代表 **Range** 对象的变量。

**说明**

将此属性设置为 **True** 以隐藏行或列。指定的区域必须占据整个行或整个列。

请不要将此属性与 **FormulaHidden** 属性混淆。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例隐藏 Sheet1 的 C 列。*/ Application.Worksheets.Item("Sheet1").Columns.Item("C").Hidden = true` |

#### **Range.HorizontalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

*express*   一个代表 **Range** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlCenter**      |
| ----------------- |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **Range.ID**

返回或设置一个 **String** 值，它代表将页面另存为网页时指定单元格的识别标志。

**语法**

**express.ID**

*express*   一个代表 **Range** 对象的变量。

**说明**

可将 ID 标签用作其他 HTML 文档或相同网页上的超链接引用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将活动工作表上 A1 单元格的 ID 设置为“target”。*/ Application.ActiveSheet.Range("A1").ID = "target"  /*之后，将文档另存为网页，并将下面的 HTML 行添加到网页中。 然后，当用户在 Web 浏览器中查看该页并用鼠标单击该超链接时，浏览器将显示该单元格。*/ <A HREF="#target">Quarterly earnings</A>` |

#### **Range.Left**

返回一个 **Variant** 值，它代表从列 A 的左边缘到区域的左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果该区域不连续，则使用第一块区域。如果该区域有若干列宽，则使用区域中最左边的列。

#### **Range.MergeArea**

返回一个 **Range** 对象，该对象代表包含指定单元格的合并区域。如果指定的单元格不在合并区域内，则该属性返回指定的单元格。只读。**Variant** 类型。

**语法**

**express.MergeArea**

*express*   一个代表 **Range** 对象的变量。

**说明**

**MergeArea** 属性只应用于单个单元格区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为包含单元格 A3 的合并区域赋值。*/ function test() {     let ma = Application.Range("a3").MergeArea     if(ma.Address == "$A$3") {         MsgBox("not merged")     }     else {         ma.Cells.Item(1, 1).Value2 = "42"     } }` |

#### **Range.MergeCells**

如果区域包含合并单元格，则为 **True**。**Variant** 型，可读写。

**语法**

**express.MergeCells**

*express*   一个代表 **Range** 对象的变量。

**说明**

选定包含合并单元格的区域时，所选定的区域可能与所期望选定的区域不同。使用 **Address** 属性检验选定区域的地址。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为包含单元格 A3 的合并区域赋值。*/ function test() {     let ma = Application.Range("a3").MergeArea     if(Application.Range("a3").MergeCells) {         MsgBox(123)         ma.Cells.Item(1, 1).Value2 = "42"     } }` |

#### **Range.Name**

返回或设置一个 **Variant** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Range** 对象的变量。

#### **Range.NumberFormat**

返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果指定区域中的所有单元格包含不同的数字格式，则此属性返回 **Null**。

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下这些示例分别对 Sheet1 中的 A17 单元格、第一行和 C 列的数字格式进行设置。*/ function test() {     Application.Worksheets.Item("Sheet1").Range("A17").NumberFormat = "General"     Application.Worksheets.Item("Sheet1").Rows.Item(1).NumberFormat = "hh:mm:ss"     Application.Worksheets.Item("Sheet1").Columns.Item("C").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)" }` |

#### **Range.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Range** 对象的变量。

#### **Range.Previous**

返回一个代表下一个单元格的 **Range** 对象。

**语法**

**express.Previous**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例选定 sheet1 中上一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格左边的单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.Previous.Select() }` |

#### **Range.Row**

返回区域中第一个子区域的第一行的行号。**Long** 类型，只读。

**语法**

**express.Row**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将工作表 sheet1 中每隔一行的行高设置为 4 磅。*/ function test() {     let rw = Application.Worksheets.Item("Sheet1").Rows     for(let i = 1; i <= rw.Count; i++) {         if((rw.Item(i).Row) % 2 == 0) {            rw.Item(i).RowHeight = 4         }     } }` |

#### **Range.RowHeight**

以磅为单位返回或设置指定区域中所有行的行高。如果指定区域中的各行的行高不等，则返回 **null**。**Variant** 类型，可读写。

**语法**

**express.RowHeight**

*express*   一个代表 **Range** 对象的变量。

**说明**

可使用 **Height** 属性返回整个单元格区域的高度。

以下是 **RowHeight** 和 **Height** 的不同之处：

- **Height** 为只读。
- 如果要返回几行的 **RowHeight** 属性，可得到每一行的行高（如果所有的行等高），或得到 **null**（如果它们不等高）。如果要返回几行的 **Height** 属性，将得到所有行高的总和。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使工作表 sheet1 上第一行的行高加倍。*/ function test() {     let ro= Application.Worksheets.Item("Sheet1").Rows.Item(1)     ro.RowHeight =ro.RowHeight * 2 }` |

#### **Range.Rows**

返回一个 **Range** 对象，它代表指定单元格区域中的行。**Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，只读。

**语法**

**express.Rows**

*express*   一个代表 **Range** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 ActiveSheet.Rows。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 Selection.Rows.Count 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 Areas.Count 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例删除 Sheet1 的第三行。*/ Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()  /*此示例检查第一张工作表上当前区域中的行，如果某行的第一个单元格值与前一行的第一个单元格的值相等，则删除此行。*/ function test() {     let rw = Application.Worksheets.Item(1).Cells.Item(1, 1).CurrentRegion.Rows     for(let i = 2; i <= rw.Count; i++) {         let ths= rw.Item(i).Cells.Item(1, 1).Value2         let last =rw.Item(i-1).Cells.Item(1, 1).Value2         if(ths == last) {             rw.Item(i).Delete()             last = ths         }     } }  /*此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let areaCount = Application.Selection.Areas.Count     if(areaCount <= 1) {         MsgBox("The selection contains " + Selection.Rows.Count & " rows.")     }     else {       let  i = 1       let a = Application.Selection.Areas       for(let b=1;b<=a.Count;b++) {           MsgBox("Area " + i + " of the selection contains " + a.Item(b).Rows.Count + " rows.")           i = i + 1       }       } }` |

#### **Range.Text**

返回或设置指定对象中的文本。**String** 型，只读。

**语法**

**express.Text**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例演示包含格式数字的单元格的 Text 和 Value 属性的区别。*/ function test() {     let c = Application.Worksheets.Item("Sheet1").Range("B14")     c.Value2 = 1198.3       c.NumberFormat = "$#,##0_);($#,##0)"     MsgBox(c.Value2)     MsgBox(c.Text) }` |

#### **Range.Top**

返回或设置一个 **Variant** 值，它代表行 1 上边缘到区域上边缘的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Range** 对象的变量。

**说明**

如果该区域不连续，则使用第一块区域。如果该区域有若干行，则使用区域中的顶端行（行号最小的行）。

#### **Range.Value2**

返回或设置单元格值。**Variant** 类型，可读写。

**语法**

**express.Value2**

*express*   一个代表 **Range** 对象的变量。

**说明**

该属性与 **Value** 属性的唯一区别是：**Value2** 属性不使用 **Currency** 和 **Date** 数据类型。可以通过使用 **Double** 数据类型，以浮点数形式返回这些数据类型格式的数值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使用 Value2 属性对两个单元格的值进行相加。*/ Application.Range("a1").Value2 = Range("b1").Value2 + Range("c1").Value2 ` |

#### **Range.Width**

返回一个 **Variant** 值，它代表区域的宽度（以单位表示）。

**语法**

**express.Width**

*express*   一个代表 **Range** 对象的变量。

#### **Range.Worksheet**

返回一个 **Worksheet** 对象，该对象表示包含指定区域的工作表。只读。

**语法**

**express.Worksheet**

*express*   一个代表 **Range** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示包含活动单元格的工作表的名称。本示例必须在工作表上运行。*/ MsgBox(Application.ActiveCell.Worksheet.Name)  /*本示例显示包含名为“testRange”的区域的工作表名称。*/ MsgBox(Application.Range("testRange").Worksheet.Name)` |

适用环境：web

适用平台：windows/linux