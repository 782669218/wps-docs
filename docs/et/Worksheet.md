**Worksheet**



代表一个工作表。

**说明**

**Worksheet** 对象是 [Worksheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheets/Worksheets%20.htm#jsObject_Worksheets)集合的成员。**Worksheets** 集合包含某个工作簿中所有的 **Worksheet** 对象。

**Worksheet** 对象也是 [Sheets ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合的成员。**Sheets** 集合包含工作簿中所有的工作表（图表工作表和工作表）。

## 示例

使用 [Worksheets](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheets/Worksheets%20.htm#jsObject_Worksheets)( *index*)（其中 *index* 是工作表索引号或名称）可返回一个 **Worksheet** 对象。下例隐藏活动工作簿中的工作表一。

| 示例代码复制                                     |
| ------------------------------------------------ |
| `Application.Worksheets.Item(1).Visible = false` |

工作表索引号指示该工作表在工作簿的标签栏上的位置。`Worksheets(1)` 是工作簿中第一个（最左边的）工作表，而 `Worksheets(Worksheets.Count)` 是最后一个。所有工作表均包括在索引计数中，即便是隐藏工作表也是如此。

工作表名称显示在该工作表的标签上。使用 [Name ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#Worksheet.Name)属性可设置或返回工作表名称。下例保护 Sheet1 上的方案。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let strPassword = prompt("Enter the password for the worksheet") Application.Worksheets.Item("Sheet1").Protect(strPassword,null,null,true)` |

当工作表处于活动状态时，可以使用ActiveSheet 属性来引用它。下例使用**Activ****ate**方法激活 Sheet1，将页面方向设置为横向，然后打印该工作表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Activate() Application.ActiveSheet.PageSetup.Orientation = xlLandscape Application.ActiveSheet.PrintOut()` |

**方法**

|                                                              | 名称          | 说明                                                         |
| ------------------------------------------------------------ | ------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**      | 将工作表复制到工作簿的另一位置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**    | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**      | 将工作表移到工作簿中的其他位置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Protect**   | 保护工作表使其不能被修改。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**     | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表一个单元格或单元格区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**    | 选择对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unprotect** | 取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Cells**       | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表工作表中的所有单元格（不仅仅是当前使用的单元格）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Columns**     | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 **Columns** 属性失效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Names**       | 返回一个 [Names ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，它代表所有特定于工作表的名称（使用“WorksheetName!”前缀定义的名称）。**Names** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**        | 返回代表下一个工作表的 [Worksheet ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**    | 返回代表下一个工作表的 [Worksheet ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Protection**  | 返回一个 [Protection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Protection/Protection%20.htm#jsObject_Protection) 对象，该对象表示工作表的保护选项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rows**        | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表指定工作表中的所有行。**Range** 对象，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**      | 返回一个 [Shapes ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合，它代表工作表上的所有形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回一个代表工作表类型的 [XlSheetType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlSheetType%20%E6%9E%9A%E4%B8%BE.html)值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsedRange**   | 返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象表示指定工作表上所使用的区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置一个 [XlSheetVisibility ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlSheetVisibility%20%E6%9E%9A%E4%B8%BE.html)值，它确定对象是否可见。 |

**成员方法**

#### **Worksheet.Copy**

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy(Before, After)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之后。*/ Application.Worksheets.Item("Sheet1").Copy(null,Application.Worksheets.Item("Sheet3"))` |

#### **Worksheet.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

在删除 **Worksheet** 时，此方法默认情况下显示一个对话框，用于提示用户确认是否删除。

#### **Worksheet.Move**

将工作表移到工作簿中的其他位置。

**语法**

**express.Move(Before, After)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 在其之前放置移动工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 在其之后放置移动工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，ET 将新建一个工作簿，其中包含所移动的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将当前活动工作簿的 Sheet1 移到 Sheet3 之后。*/ Application.Worksheets.Item("Sheet1").Move(null,Application.Worksheets.Item("Sheet3"))` |

#### **Worksheet.Protect**

保护工作表使其不能被修改。

**语法**

**express.Protect(Password, DrawingObjects, Contents, Scenarios, UserInterfaceOnly, AllowFormattingCells, AllowFormattingColumns, AllowFormattingRows, AllowInsertingColumns, AllowInsertingRows, AllowInsertingHyperlinks, AllowDeletingColumns, AllowDeletingRows, AllowSorting, AllowFiltering, AllowUsingPivotTables)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称**                   | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Password*                 | 可选          | **Variant**  | 一个字符串，该字符串为工作表或工作簿指定区分大小写的密码。如果省略此参数，不用密码就可以取消对工作表或工作簿的保护。否则，必须指定密码才能取消对工作表或工作簿的保护。如果忘记了密码，就无法取消对工作表或工作簿的保护。 |
| *DrawingObjects*           | 可选          | **Variant**  | 如果为 True，则保护形状。默认值是 True。                     |
| *Contents*                 | 可选          | **Variant**  | 如果为 True，则保护内容。对于图表，这样会保护整个图表。对于工作表，这样会保护锁定的单元格。默认值是 True。 |
| *Scenarios*                | 可选          | **Variant**  | 如果为 True，则保护方案。此参数仅对工作表有效。默认值是 True。 |
| *UserInterfaceOnly*        | 可选          | **Variant**  | 如果为 True，则保护用户界面，但不保护宏。如果省略此参数，则既保护宏也保护用户界面。 |
| *AllowFormattingCells*     | 可选          | **Variant**  | 如果为 True，则允许用户为受保护的工作表上的任意单元格设置格式。默认值是 False。 |
| *AllowFormattingColumns*   | 可选          | **Variant**  | 如果为 True，则允许用户为受保护的工作表上的任意列设置格式。默认值是 False。 |
| *AllowFormattingRows*      | 可选          | **Variant**  | 如果为 True，则允许用户为受保护的工作表上的任意行设置格式。默认值是 False。 |
| *AllowInsertingColumns*    | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上插入列。默认值是 False。 |
| *AllowInsertingRows*       | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上插入行。默认值是 False。 |
| *AllowInsertingHyperlinks* | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表中插入超链接。默认值是 False。 |
| *AllowDeletingColumns*     | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上删除列，要删除的列中的每个单元格都被解除锁定。默认值是 False。 |
| *AllowDeletingRows*        | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上删除行，要删除的行中的每个单元格都被解除锁定。默认值是 False。 |
| *AllowSorting*             | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上进行排序。排序区域中的每个单元格必须是解除锁定的或取消保护的。默认值是 False。 |
| *AllowFiltering*           | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上设置筛选。用户可以更改筛选条件，但是不能启用或禁用自动筛选功能。用户也可以在已有的自动筛选功能上设置筛选。默认值是 False。 |
| *AllowUsingPivotTables*    | 可选          | **Variant**  | 如果为 True，则允许用户在受保护的工作表上使用数据透视表。默认值是 False。 |

**说明**

要在受保护的工作表上做更改，如果提供密码，则可在受保护的工作表上使用 **Protect** 方法。另一种方法是：取消工作表保护，对工作表做一些必要的更改，然后再次保护工作表。

**注释：**“取消保护”的意思是单元格可以被锁定（“设置单元格格式”对话框），但在“允许用户编辑区域”对话框中定义的单元格区域内，并且用户已通过密码取消了对单元格区域的保护或已通过 NT 权限的验证。

#### **Worksheet.Range**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表一个单元格或单元格区域。

**语法**

**express.Range(Cell1, Cell2)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cell1*  | 必选          | **Variant**  | 区域名称。必须为采用宏语言的 A1 样式引用。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可包括货币符号，但它们会被忽略掉。您可以在区域中任一部分使用局部定义名称。如果使用名称，则假定该名称使用的是宏语言。 |
| *Cell2*  | 可选          | **Variant**  | 区域左上角和右下角的单元格。可以是一个包含单个单元格、整列或整行的 Range 对象，或者也可以是一个用宏语言为单个单元格命名的字符串。 |

**说明**

如果在没有对象识别符时使用，则该属性是 `ActiveSheet.Range` 的快捷方式（它返回活动表的一个区域，如果活动表不是一张工作表，则该属性无效）。

当应用于 **Range** 对象时，该属性与 **Range** 对象相关。例如，如果选中单元格 C3，那么 `Selection.Range("B1")` 返回单元格 D3，因为它同 **Selection** 属性返回的 **Range** 对象相关。此外，代码 `ActiveSheet.Range("B1")` 总是返回单元格 B1。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将 Sheet1 上 A1 单元格的值设置为 3.14159。*/ Application.Worksheets.Item("Sheet1").Range("A1").Value2 = 3.14159  /*此示例在 Sheet1 的 A1 单元格中创建一个公式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Formula = "=10*RAND()"  /*此示例在 Sheet1 上的单元格区域 A1:D10 中进行循环。如果某个单元格的值小于 0.001，则此代码将用 0（零）来取代该值。*/ function test() {     for(let i = 1; i <= Application.Worksheets.Item("Sheet1").Range("A1:D10").Count; i++){         if(Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 < .001 ){             Application.Worksheets.Item("Sheet1").Range("A1:D10").Item(i).Value2 = 0         }     } }  /*此示例在名为“TestRange”的区域上进行循环，并显示该区域中空白单元格的个数。*/ function test() {     let numBlanks = 0     for(let i = 1; i <= Application.Range("TestRange").Count; i++){         if(Application.Range("TestRange").Item(i).Value2 == null){             numBlanks++         }     }     alert("There are " + numBlanks + " empty cells in this range") }  /*此示例将 Sheet1 中单元格区域 A1:C5 上的字体样式设置为斜体。此示例使用 Range 属性的语法 2。*/ Application.Worksheets.Item("Sheet1").Range(Application.Cells.Item(1, 1), Application.Cells.Item(5, 3)).Font.Italic = true` |

#### **Worksheet.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**说明**

要选择单元格或单元格区域，请使用 **Select** 方法。要使单个单元格成为活动单元格，请使用**Activate**方法。

#### **Worksheet.Unprotect**

取消工作表或工作簿的保护。如果工作表或工作簿不是受保护的，则此方法不起作用。

**语法**

**express.Unprotect(Password)**

*express*   一个代表 **Worksheet** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Password* | 可选          | **Variant**  | 一个字符串，指定用于解除工作表或工作簿保护的密码，此密码是区分大小写的。如果工作表或工作簿不设密码保护，则省略此参数。如果对工作表省略此参数，而该工作表又设有密码保护，ET 将提示您输入密码。如果对工作簿省略此参数，而该工作簿又设有密码保护，则该方法将失效。 |

**说明**

如果您忘记了密码，将不能取消工作表或工作簿的保护。建议将密码和对应文档名妥善保存。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例取消当前活动工作表的保护。*/ Application.ActiveSheet.Unprotect()` |

**成员属性**

#### **Worksheet.Application**

如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }else {         alert("This is not an ET Application object.")     } }  ` |

#### **Worksheet.Cells**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表工作表中的所有单元格（不仅仅是当前使用的单元格）。

**语法**

**express.Cells**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

因为**Item**属性是 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象的默认属性，所以可以在 **Cells** 关键字后面紧接着指定行和列索引。有关详细信息，请参阅 **Item** 属性和本主题的示例。

在不使用对象识别符的情况下，使用此属性将返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中所有的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 中单元格 C5 的字号设置为 14 磅。*/ Application.Worksheets.Item("Sheet1").Cells.Item(5, 3).Font.Size = 14  /*本示例清除 Sheet1 上第一个单元格的公式。*/ Application.Worksheets.Item("Sheet1").Cells.Item(1).ClearContents()  /*本示例将 Sheet1 上所有单元格的字体设置为 8 磅“Arial”字体*/ function test() {     Application.Worksheets.Item("Sheet1").Cells.Font.Name = "Arial"     Application.Worksheets.Item("Sheet1").Cells.Font.Size = 8 }` |

#### **Worksheet.Columns**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表活动工作表中的所有列。如果活动文档不是工作表，则 **Columns** 属性失效。

**语法**

**express.Columns**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveSheet.Columns`。

此属性在应用于一个是多重选定区域的 **Range** 对象时，会只从该区域的第一个子区域中返回列。例如，如果 **Range** 对象有两个子区域 A1:B2 和 C3:D4，那么，`Selection.Columns.Count` 的返回值是 2，而不是 4。若要对一个可能包含多重选定区域的区域使用此属性，请测试 `Areas.Count` 以确定此区域内是否包含多个子区域。如果包含，请对此区域内的每个子区域进行循环。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将 Sheet1 的第一列（A 列）的字体设置为加粗。*/ Application.Worksheets.Item("Sheet1").Columns.Item(1).Font.Bold = true` |

#### **Worksheet.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

#### **Worksheet.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Worksheet** 对象的变量。

#### **Worksheet.Names**

返回一个 [Names ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#jsObject_Names)集合，它代表所有特定于工作表的名称（使用“WorksheetName!”前缀定义的名称）。**Names** 对象，只读。

**语法**

**express.Names**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveWorkbook.Names`。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例是将 Sheet1 中的 A1 单元格的名称定义为“myName”。*/ Application.ActiveWorkbook.Names.Add("myName", null, null, null, null, null, null, null, null, "=Sheet1!R1C1")` |

#### **Worksheet.Next**

返回代表下一个工作表的 [Worksheet ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。

**语法**

**express.Next**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

如果该对象为区域，则此属性会模拟 Tab 键，但会返回下一个单元格，而不选中它。

在处于保护状态的工作表中，此属性返回下一个未锁定单元格。在未保护的工作表中，此属性总是返回紧靠指定单元格右边的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例选定 sheet1 中下一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格右边的单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.Next.Select() }` |

#### **Worksheet.Previous**

返回代表下一个工作表的 [Worksheet ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。

**语法**

**express.Previous**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

如果指定对象为区域，则此属性的作用是仿效 Shift+Tab；但此属性只是返回上一单元格，并不选定它。

在保护工作表上，该属性返回上一个未锁定的单元格。在未保护的工作表上，该属性通常返回指定单元格左侧相邻的单元格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例选定 sheet1 中上一个未锁定单元格。如果 sheet1 未保护，选定的单元格将是紧靠活动单元格左边的单元格。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.Previous.Select() }` |

#### **Worksheet.Protection**

返回一个 [Protection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Protection/Protection%20.htm#jsObject_Protection) 对象，该对象表示工作表的保护选项。

**语法**

**express.Protection**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动工作表进行保护，并判断是否能在受保护的工作表中插入列，然后将此状态通知用户。*/ function test() {     Application.ActiveSheet.Protect()      // Check the ability to insert columns on a protected sheet.     // Notify the user of this status.     if(Application.ActiveSheet.Protection.AllowInsertingColumns == true){         alert("The insertion of columns is allowed on this protected worksheet.")     }     else{         alert("The insertion of columns is not allowed on this protected worksheet.")     } }` |

#### **Worksheet.Rows**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，它代表指定工作表中的所有行。**Range** 对象，只读。

**语法**

**express.Rows**

*express*   一个代表 **Worksheet** 对象的变量。

**说明**

在不使用对象识别符的情况下使用此属性等效于使用 `ActiveSheet.Rows`。

该属性在应用于是多个选定区域的 **Range** 对象时，只从该区域中第一个子区域内返回行。例如，如果 **Range** 对象有两个子区域：A1:B2 和 C3:D4，则 `Selection.Rows.Count` 返回 2 而不是 4。若要在一个可能包含多个选定区域的区域中使用此属性，请测试 `Areas.Count` 以确定该区域是否包含多个选择区域。如果是，请对此区域内的每个子区域进行循环，如第 3 个示例所示。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例删除 Sheet1 的第三行。*/ Application.Worksheets.Item("Sheet1").Rows.Item(3).Delete()  /*此示例显示 Sheet1 选定区域中的行数。如果选择了多个子区域，此示例将对每一个子区域进行循环。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     let areaCount = Application.Selection.Areas.Count     if(areaCount <= 1){        alert("The selection contains " + Application.Selection.Rows.Count + " rows.")     }     else{         for(let i = 1; i <= Application.Selection.Areas.Count; i++){             alert("Area " + i + " of the selection contains " + Application.Selection.Areas.Rows.Count + " rows.")         }     } }` |

#### **Worksheet.Shapes**

返回一个 [Shapes ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Shapes/Shapes%20.htm#jsObject_Shapes)集合，它代表工作表上的所有形状。只读。

**语法**

**express.Shapes**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例向第一张工作表中添加兰色的虚线。*/ function test() {     let Wshapes = Application.Worksheets.Item(1).Shapes.AddLine(10, 10, 250, 250).Line     Wshapes.DashStyle = Application.Enum.msoLineDashDotDot     Wshapes.ForeColor.RGB = (50, 0, 128) }` |

#### **Worksheet.Type**

返回一个代表工作表类型的 [XlSheetType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlSheetType%20%E6%9E%9A%E4%B8%BE.html)值。

**语法**

**express.Type**

*express*   一个代表 **Worksheet** 对象的变量。

#### **Worksheet.UsedRange**

返回一个 [Range ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象表示指定工作表上所使用的区域。只读。

**语法**

**express.UsedRange**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定 Sheet1 中的已用区域。*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveSheet.UsedRange.Select() }` |

#### **Worksheet.Visible**

返回或设置一个 [XlSheetVisibility ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlSheetVisibility%20%E6%9E%9A%E4%B8%BE.html)值，它确定对象是否可见。

**语法**

**express.Visible**

*express*   一个代表 **Worksheet** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏 Sheet1。*/ Application.Worksheets.Item("Sheet1").Visible = false  /*本示例将 Sheet1 设置为可见。*/ Application.Worksheets.Item("Sheet1").Visible = true  /*本示例使活动工作簿中的每一张工作表可见。*/ for(let i = 1; i <= Sheets.Count; i++){     Application.Sheets.Visible = true }  /*本示例新建一张工作表，然后将其 Visible 属性设为 xlVeryHidden。要引用该工作表，可使用其对象变量 newSheet，如本示例最后一行所示。若要在另一过程中使用 newSheet 对象变量，必须在模块的第一行中将其声明为公用变量 (Public newSheet As Object)，然后才能执行 Sub 或 Function 过程。*/ function test() {     let newSheet = Application.Worksheets.Add()     newSheet.Visible = Application.Enum.xlVeryHidden     newSheet.Range("A1:D4").Formula = "=RAND()" }` |

适用环境：web

适用平台：windows/linux