**Scenarios**



指定工作表上所有 **Scenario** 对象的集合。

**说明**

方案是一组经过命名和保存的输入值（被称为“可变单元格”）。

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**           | 新建方案并将其添加到当前工作表可用的方案列表中。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateSummary** | 新建一个工作表，并在其中显示指定工作表中所有方案的摘要报表。**Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**          | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge**         | 将另一张工作表中的方案合并到 **Scenarios** 集合中。          |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Scenarios.Add**

新建方案并将其添加到当前工作表可用的方案列表中。

**语法**

**express.Add(Name, ChangingCells, Values, Comment, Locked, Hidden)**

*express*   一个代表 **Scenarios** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*          | 必选          | **String**   | 方案名。                                                     |
| *ChangingCells* | 必选          | **Variant**  | 一个 Range 对象，它引用方案的可变单元格。                    |
| *Values*        | 可选          | **Variant**  | 包含 ChangingCells 中单元格方案值的数组。如果省略此参数，就假定方案值是 ChangingCells 中单元格的当前值。 |
| *Comment*       | 可选          | **Variant**  | 一个指定方案批注文本的字符串。添加新方案时，作者的姓名和日期自动添加在批注文本的开始部分。 |
| *Locked*        | 可选          | **Variant**  | 如果为 True，则锁定方案以防更改。默认值为 True。             |
| *Hidden*        | 可选          | **Variant**  | 如果为 True，则隐藏方案。默认值为 False。                    |

**返回值**

一个代表新方案的 Scenario 对象。

**说明**

方案名称必须是唯一的；如果试图用已经在使用的名称创建方案，ET 会产生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向工作表 Sheet1 添加新方案。 Application.Worksheets.Item("Sheet1").Scenarios.Add("Best Case",Worksheets("Sheet1").Range("A1:A4"), [23, 5, 6, 21],"Most favorable outcome.")` |

#### **Scenarios.CreateSummary**

新建一个工作表，并在其中显示指定工作表中所有方案的摘要报表。**Variant** 类型。

**语法**

**express.CreateSummary(ReportType, ResultCells)**

*express*   一个代表 **Scenarios** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**            | **说明**                                                     |
| ------------- | ------------- | ----------------------- | ------------------------------------------------------------ |
| *ReportType*  | 可选          | **XlSummaryReportType** | 指定摘要报表是数据透视表还是标准摘要。                       |
| *ResultCells* | 可选          | **Variant**             | 一个代表指定工作表中计算结果单元格的 Range 对象。一般而言，此区域指一个或多个包含公式的单元格，即显示特定方案计算结果的单元格；这些单元格中的公式的值由模型中的可变单元格值计算而得。如果省略此参数，报表中将没有任何结果单元格。 |

**返回值**

Variant

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例用工作表 Sheet1 上单元格区域 C4:C9 的计算结果创建 Sheet1 上方案的摘要报表。 Application.Worksheets.Item("Sheet1").Scenarios.CreateSummary(undefined, Worksheets.Item("Sheet1").Range("C4:C9"))` |

#### **Scenarios.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Scenarios** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Scenario 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为“Options”的工作表上显示名为“Typical”的方案。 Application.Worksheets.Item("options").Scenarios.Item("typical").Show()` |

#### **Scenarios.Merge**

将另一张工作表中的方案合并到 **Scenarios** 集合中。

**语法**

**express.Merge(Source)**

*express*   一个代表 **Scenarios** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Source* | 必选          | **Variant**  | 待合并方案所在工作表的名称，或代表该工作表的 Worksheet 对象。 |

**返回值**

Variant

**说明**

合并区域的值在该区域左上角的单元格中指定。

**成员属性**

#### **Scenarios.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Scenarios** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET")     {         alert("This is an ET Application object.")     }     else     {         alert("This is not an ET Application object.")     } }` |

#### **Scenarios.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Scenarios** 对象的变量。

#### **Scenarios.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Scenarios** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Scenarios.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Scenarios** 对象的变量。

适用环境：web

适用平台：windows/linux