**Name**



代表单元格区域的定义名。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

**说明**

应用程序、工作簿和 **Worksheet** 对象

**Name** 对象是 **Application**、**Workbook** 和 **Worksheet** 对象的 **Names** 集合的成员。使用 **Names**(*index* )（其中 *index* 是名称索引号或定义名称）可返回一个 **Name** 对象。

索引号表明名称在集合中的位置。名称按字母顺序从 a 到 z 放置，不区分大小写。

**Range 对象**

虽然 **Range** 对象可以有多个名称，但 **Range** 对象没有 **Names** 集合。将 **Name** 与一个 **Range** 对象一起使用可从名称列表（按字母顺序排序）中返回第一个名称。下例为指定给工作表一上单元格 A1:B1 的第一个名称设置 **Visible** 属性。

## 示例

本示例显示应用程序集合中第一个名称的单元格引用。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     alert(Application.Application.Names.Item(1).RefersTo)  } ` |

下例从活动工作簿中删除名称“mySortRange”。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveWorkbook.Names.Item("mySortRange").Delete() } ` |

使用 

Name

 属性可返回或设置名称本身的文本。本示例更改活动工作簿中第一个 

Name

 对象的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Names.Item(1).Name = "stock_values" } ` |

下例为指定给工作表一上单元格 A1:B1 的第一个名称设置 

Visible

 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item(1).Range("a1:b1").Name.Visible = false } ` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Category**      | 返回或设置指定名称在宏语言中的分类。该名称必须针对一个自定义函数或命令。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comment**       | 返回或设置与名称关联的批注。可读/写 **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**         | 返回 **Number** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**          | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefersTo**      | 用宏语言以 A1 样式表示法返回或设置名称所引用的公式（以等号开头）。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefersToR1C1**  | 返回或设置指定名称所引用的公式。公式以等号开头，由宏语言和 R1C1-样式引用组成。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefersToRange** | 返回由 **Name** 对象引用的 **Range** 对象。只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**         | 返回或设置一个 **String** 值，它代表规定名称去引用的公式。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**       | 返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。  |

**成员方法**

#### **Name.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Name** 对象的变量。

**成员属性**

#### **Name.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Name** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */ function test() {     let myObject = Application.ActiveWorkbook     if (myObject.Application.Value == "Microsoft Excel") {         alert("This is a Microsoft Excel Application object.")     }     else {         alert("This is not a Microsoft Excel Application object.")     } }` |

#### **Name.Category**

返回或设置指定名称在宏语言中的分类。该名称必须针对一个自定义函数或命令。

String

 类型，可读写。

**语法**

**express.Category**

*express*   一个代表 **Name** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例假定已在 ET 4.0 宏表中创建了一个自定义函数或命令。本示例显示该名称在宏语言中的分类。假定该自定义函数或命令的名称在指定工作簿中唯一。*/ function test() {     if (Application.ActiveWorkbook.Names.Item('test').MacroType != Application.Enum.xlNone) {         alert("The category for this name is " + Application.ActiveWorkbook.Names.Item('test').Category)     }     else {         alert("This name does not refer to a custom function or command.")     } }` |

#### **Name.Comment**

返回或设置与名称关联的批注。可读/写 **String** 类型。

**语法**

**express.Comment**

*express*   一个代表 **Name** 对象的变量。

**说明**

批注文本不能超过 255 个字符。

#### **Name.Index**

返回 **Number** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Name** 对象的变量。

#### **Name.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Name** 对象的变量。

#### **Name.RefersTo**

用宏语言以 A1 样式表示法返回或设置名称所引用的公式（以等号开头）。**String** 类型，可读写。

**语法**

**express.RefersTo**

*express*   一个代表 **Name** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例创建活动工作簿中所有名称的列表，并用宏语言以 A1 样式表示法表示这些名称所引用的公式。该列表将出现在本示例新建的工作表中。 */ function test() {     let newSheet = Application.Worksheets.Add()     let i = 1     for (let nm = 1; nm <= Application.ActiveWorkbook.Names.Count; nm++) {         newSheet.Cells.Item(i, 1).Value2 = Application.ActiveWorkbook.Names.Item(nm).Name         newSheet.Cells.Item(i, 2).Value2 = "'" + Application.ActiveWorkbook.Names.Item(nm).RefersTo         i = i + 1     }     newSheet.Columns.Item("A:B").AutoFit() }` |

#### **Name.RefersToR1C1**

返回或设置指定名称所引用的公式。公式以等号开头，由宏语言和 R1C1-样式引用组成。**String** 类型，可读写。

**语法**

**express.RefersToR1C1**

*express*   一个代表 **Name** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例新建一个工作表，并将当前工作簿所有名称的列表插入到新工作表中，名称列表中包括其公式（由 R1C1-样式引用和宏语言组成）。 */ function test() {     let newSheet = Application.ActiveWorkbook.Worksheets.Add()     let i = 1     for (let nm = 1; nm <= Application.ActiveWorkbook.Names.Count; nm++) {         newSheet.Cells.Item(i, 1).Value2 = Application.ActiveWorkbook.Names.Item(nm).Name         newSheet.Cells.Item(i, 2).Value2 = "'" + Application.ActiveWorkbook.Names.Item(nm).RefersToR1C1         i = i + 1     } }` |

#### **Name.RefersToRange**

返回由 **Name** 对象引用的 **Range** 对象。只读。

**语法**

**express.RefersToRange**

*express*   一个代表 **Name** 对象的变量。

**说明**

如果 **Name** 对象并不引用区域（例如，如果该对象引用的是一个常量或公式），则该属性无效。

要更改名称所引用的区域，请使用 **RefersTo** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示当前工作表打印区域中的行数和列数。 */ /* 注释:确保在当前工作簿的活动工作表上建立打印区域。 */ function test() {     let p = Application.Sheets.Item(Application.ActiveSheet.Name).Names.Item("Print_Area").RefersToRange.Value2     alert("Print_Area: " + p.length + " rows, " + p[0].length + " columns") } ` |

#### **Name.Value**

返回或设置一个 **String** 值，它代表规定名称去引用的公式。

**语法**

**express.Value**

*express*   一个代表 **Name** 对象的变量。

#### **Name.Visible**

返回或设置一个 **Boolean** 值，它确定对象是否可见。可读写。

**语法**

**express.Visible**

*express*   一个代表 **Name** 对象的变量。

适用环境：web

适用平台：windows/linux