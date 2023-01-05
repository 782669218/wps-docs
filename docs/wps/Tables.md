#### **Tables**



**Table** 对象的集合，这些对象代表所选内容、范围或文档中的表格。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **Table** 对象，该对象代表添加到文档中的新的空白表格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Table** 对象。                            |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**        | 返回一个 **Long** 类型的值，该值代表集合中表格的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NestingLevel** | 返回指定表格的嵌套层。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回一个 **Object** 类型值，该值代表指定 **Tables** 对象的父对象。 |

**成员方法**

#### **Tables.Add**

返回一个 **Table** 对象，该对象代表添加到文档中的新的空白表格。

**语法**

**express.Add(Range, NumRows, NumColumns, DefaultTableBehavior , AutoFitBehavior)**

*express*   一个代表 **Tables** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**     | **说明**                                                     |
| ---------------------- | ------------- | ---------------- | ------------------------------------------------------------ |
| *Range*                | 必选          | **Range object** | 表格出现的区域。如果该区域未折叠，表格将替换该区域。         |
| *NumRows*              | 必选          | **Long**         | 要在表格中包括的行数。                                       |
| *NumColumns*           | 必选          | **Long**         | 要在表格中包括的列数。                                       |
| *DefaultTableBehavior* | 可选          | **Variant**      | 设置一个值来指定 WPS 是否要根据单元格的内容自动调整表格单元格的大小（“自动调整”功能）。可以是下列常量之一：**wdWord8TableBehavior**（禁用“自动调整”功能）或 **wdWord9TableBehavior**（启用“自动调整”功能）。默认常量为 **wdWord8TableBehavior**。 |
| *AutoFitBehavior*      | 可选          | **Variant**      | 用于设置 WPS 调整表格大小的“自动调整”规则。可以为一个 **WdAutoFitBehavior** 常量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头添加一个 3 行 4 列的空表格*/ let myRange = Application.ActiveDocument.Range(0, 0) Application.ActiveDocument.Tables.Add(myRange, 3, 4)  /*本示例在活动文档的结尾添加一个 6 行 10 列的空白新表格*/ let myRange = Application.ActiveDocument.Content myRange.Collapse(wdCollapseEnd) ActiveDocument.Tables.Add(myRange, 6, 10)  /*本示例向新文档中添加一个 3 行 5 列的表格，然后在表格的每个单元格中插入数据*/ function NewTable(){     let docNew = Application.Documents.Add()     let tblNew = docNew.Tables.Add(Selection.Range, 3, 5)     for(let i = 1; i <= 3; i++){         for(let j=1; j <=5; j++){             tblNew.Cell(i, j).Range.InsertAfter("Cell: R" + i + ", C" + j)         }     }     tblNew.Columns.AutoFit() }` |

#### **Tables.Item**

返回集合中的单个 **Table** 对象。

**语法**

**express.Item(Index )**

*express*   一个代表 **Tables** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 **Long** 类型值。 |

**成员属性**

#### **Tables.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Tables** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Tables.Count**

返回一个 **Long** 类型的值，该值代表集合中表格的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **Tables** 对象的变量。

#### **Tables.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Tables** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Tables.NestingLevel**

返回指定表格的嵌套层。只读 **Long** 类型。

**语法**

**express.NestingLevel**

*express*   一个代表 **Tables** 对象的变量。

**说明**

最外围表格的嵌套层为 1。每一个相连嵌套表格的嵌套层比其前面表格的嵌套层高 1。

#### **Tables.Parent**

返回一个 **Object** 类型值，该值代表指定 **Tables** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Tables** 对象的变量。

适用环境：web

适用平台：windows/linux