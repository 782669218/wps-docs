**Names**



应用程序或工作簿中所有 **Name** 对象的集合。

**说明**

每一个 **Name** 对象都代表一个单元格区域的定义名称。名称可以是内置名称（如“Database”、“Print_Area”和“Auto_Open”）或自定义名称。

*RefersTo* 参数必须以 A1 样式表示法指定，包括必要时使用的美元符 ($)。例如，如果在 Sheet1 上选定了单元格 A10，并且通过将 *RefersTo* 参数“=Sheet1!A1:B1”而定义了一个名称，那么该新名称实际上指向单元格区域 A10:B10（因为指定的是相对引用）。若要指定绝对引用，请使用“=Sheet1!$A$1:$B$1”。

## 示例

使用 **Names** 属性可返回 **Names** 集合。下例创建活动工作簿中所有名称及其引用地址的列表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let nms = Application.ActiveWorkbook.Names     let wks = Application.Worksheets.Item(1)     for(let r = 1; r <= nms.Count; r++){         wks.Cells.Item(r, 2).Value2 = nms.Item(r).Name         wks.Cells.Item(r, 3).Value2 = nms.Item(r).RefersToRange.Address     } } ` |

使用 

[Add](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#Names.Add)

[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Names/Names%20.htm#Names.Add)方法可创建一个名称并将它添加到集合。下例创建一个新名称，该名称引用名为“Sheet1”的工作表上的单元格 A1:C20。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Names.Add ("test1", "=sheet1!$a$1:$c$20") } ` |

使用 

Names

(

index

)（其中 

index

 是名称索引号或定义名称）可返回一个 

Name

 对象。下例从活动工作簿中删除名称“mySortRange”。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveWorkbook.Names.Item("mySortRange").Delete() } ` |

**方法**

|                                                              | 名称     | 说明                                      |
| ------------------------------------------------------------ | -------- | ----------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 为单元格区域定义新名称。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从 **Names** 集合返回一个 **Name** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Number** 值，它代表集合中对象的数量。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Names.Add**

为单元格区域定义新名称。

**语法**

**express.Add(Name, RefersTo, Visible, MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal, RefersToR1C1, RefersToR1C1Local)**

*express*   一个代表 **Names** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*              | 可选          | **Variant**  | 如果未指定 NameLocal 参数，则指定要用作名称的英文文本。名称不能包括空格，并且不能设置为单元格引用的格式。 |
| *RefersTo*          | 可选          | **Variant**  | 如果未指定 RefersToLocal、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以英文表示）。 注释 如果引用不存在，则返回 Nothing。 |
| *Visible*           | 可选          | **Variant**  | True 指定将名称定义为可见。False 指定将名称定义为隐藏。已隐藏的名称不会在“定义名称”、“粘贴名称”或“转到”对话框中显示。默认值为 True。 |
| *MacroType*         | 可选          | **Variant**  | 由以下值之一确定的宏类型： 1 - 用户定义函数（Function 过程） 2 - 宏（Sub 过程） 3 或省略 - 无（该名称不引用用户定义函数或宏） |
| *ShortcutKey*       | 可选          | **Variant**  | 指定宏的快捷键。必须是单个字母，例如“z”或“Z”。仅适用于命令宏。 |
| *Category*          | 可选          | **Variant**  | 如果 MacroType 参数等于 1 或 2，则此参数为宏或函数的分类。该分类在“函数向导”中使用。可以用数字（从 1 开始）或名称（以英文指定）引用现有的分类。如果指定的分类不存在，ET 将创建新分类。 |
| *NameLocal*         | 可选          | **Variant**  | 如果未指定 Name 参数，则指定要用作名称的本地化的文本。名称不能包括空格，并且不能设置为单元格引用的格式。 |
| *RefersToLocal*     | 可选          | **Variant**  | 如果未指定 RefersTo、RefersToR1C1 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 A1 格式表示法以本地化的文本表示）。 |
| *CategoryLocal*     | 可选          | **Variant**  | 如果未指定 Category 参数，则指定标识自定义函数分类的本地化的文本。 |
| *RefersToR1C1*      | 可选          | **Variant**  | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1Local 参数，则说明名称引用的内容（使用 R1C1 格式表示法以英文表示）。 |
| *RefersToR1C1Local* | 可选          | **Variant**  | 如果未指定 RefersTo、RefersToLocal 和 RefersToR1C1 参数，则说明名称引用的内容（使用 R1C1 格式表示法以本地化的文本表示）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例为活动工作簿中 Sheet1 上的区域 A1:D3 定义一个新名称。 */ /* 注释:如果工作表 Sheet1 不存在，则返回 Nothing。 */ function test() {     Application.ActiveWorkbook.Names.Add("tempRange", "=Sheet1!$A$1:$D$3") } ` |

#### **Names.Item**

从 **Names** 集合返回一个 **Name** 对象。

**语法**

**express.Item(Index, IndexLocal, RefersTo)**

*express*   一个代表 **Names** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*      | 可选          | **Variant**  | 要返回的定义名称的名称或编号。                               |
| *IndexLocal* | 可选          | **Variant**  | 以用户语言表示的定义名称的名称。如果使用该参数，将不转换名称。 |
| *RefersTo*   | 可选          | **Variant**  | 名称引用的内容。使用该参数以引用的内容标识名称。             |

**说明**

必须只指定这三个参数中的一个。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例将活动工作簿中的 mySortRange 名称删除。 */ function test() {     Application.ActiveWorkbook.Names.Item("mySortRange").Delete() }` |

**成员属性**

#### **Names.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Names** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息。 */ function test() {     let myObject = Application.ActiveWorkbook     if (myObject.Application.Value == "Microsoft Excel") {         alert("This is a Microsoft Excel Application object.")     }     else {         alert("This is not a Microsoft Excel Application object.")     } }` |

#### **Names.Count**

返回一个 **Number** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Names** 对象的变量。

#### **Names.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Names** 对象的变量。

适用环境：web

适用平台：windows/linux