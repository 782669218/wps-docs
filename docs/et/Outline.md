**Outline**



代表工作表上的分级显示。

**方法**

|                                                              | 名称           | 说明                                      |
| ------------------------------------------------------------ | -------------- | ----------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowLevels** | 显示指定行号和/或列号在分级显示中的层次。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutomaticStyles** | 如果分级显示使用自动样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SummaryColumn**   | 返回或设置分级显示中的汇总列的位置。**XlSummaryColumn** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SummaryRow**      | 返回或设置分级显示中汇总行的位置。**XlSummaryRow** 类型，可读写。 |

**成员方法**

#### **Outline.ShowLevels**

显示指定行号和/或列号在分级显示中的层次。

**语法**

**express.ShowLevels(RowLevels, ColumnLevels)**

*express*   一个代表 **Outline** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *RowLevels*    | 可选          | **Variant**  | 指定分级显示中的行层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对行采取任何操作。 |
| *ColumnLevels* | 可选          | **Variant**  | 指定分级显示中的列层次。如果该分级显示包含的层次数少于指定层次，则 ET 显示所有层次。如果该参数为 0（零）或者省略该参数，则不对列采取任何操作。 |

**返回值**

Variant

**说明**

必须至少指定一个参数。

**示例**

本示例对工作表 Sheet1 的第一行到第三行的行层次和第一列的列层次进行分级显示。

| 示例代码复制                                          |
| ----------------------------------------------------- |
| `Worksheets.Item("Sheet1").Outline.ShowLevels (3, 1)` |

**成员属性**

#### **Outline.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Outline** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook        if (myObject.Application.Value == "ET"){ 　　　　    MsgBox ("This is an ET Application object.") 　　　　} 　　　　else{ 　　　　    MsgBox ("This is not an ET Application object.") 　　　　} }` |

#### **Outline.AutomaticStyles**

如果分级显示使用自动样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutomaticStyles**

*express*   一个代表 **Outline** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sheet1 的分级显示设置为使用自动样式。*/ function test(){ 　　　　Worksheets.Item("Sheet1").Outline.AutomaticStyles = true }` |

#### **Outline.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Outline** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Outline.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Outline** 对象的变量。

#### **Outline.SummaryColumn**

返回或设置分级显示中的汇总列的位置。**XlSummaryColumn** 类型，可读写。

**语法**

**express.SummaryColumn**

*express*   一个代表 **Outline** 对象的变量。

**说明**

| XlSummaryColumn 可为以下 常量之一。                       |
| --------------------------------------------------------- |
| xlSummaryOnRight 分级显示中的汇总列位于明细数据列的右侧。 |
| xlSummaryOnLeft 分级显示中的汇总列位于明细数据列的左侧。  |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个使用自动样式的分级显示，汇总行位于明细数据行的上方，汇总列位于明细数据列的右侧。*/ function test(){ 　　　　Worksheets.Item("Sheet1").Activate() 　　　　Selection.AutoOutline() 　　　　let ole = ActiveSheet.Outline 　　　　ole.SummaryRow = xlAbove 　　　　ole.SummaryColumn = xlRight 　　　　ole.AutomaticStyles = true }` |

#### **Outline.SummaryRow**

返回或设置分级显示中汇总行的位置。**XlSummaryRow** 类型，可读写。

**语法**

**express.SummaryRow**

*express*   一个代表 **Outline** 对象的变量。

**说明**

| XlSummaryRow 可为以下常量之一。                         |
| ------------------------------------------------------- |
| xlSummaryBelow 分级显示中的汇总行位于明细数据行的下方。 |
| xlSummaryAbove 分级显示中的汇总行位于明细数据行的上方。 |

若要进行 WPS 风格的分级显示，请将 **XlSummaryRow** 设置为 **xlAbove**，使分类标题位于明细信息的上方。若要进行会计风格的分级显示，请将 **XlSummaryRow** 设置为 **xlBelow**，使求和位于明细信息的下方。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个使用自动样式的分级显示，汇总行位于明细数据行的上方，汇总列位于明细数据列的右侧。*/ function test(){ 　　　　Worksheets.Item("Sheet1").Activate() 　　　　Selection.AutoOutline() 　　　　let ole = ActiveSheet.Outline 　　　　ole.SummaryRow = xlAbove 　　　　ole.SummaryColumn = xlRight 　　　　ole.AutomaticStyles = true }` |

适用环境：web

适用平台：windows/linux