**FormatConditions**



代表一个区域内所有条件格式的集合。

**说明**

**FormatConditions** 集合可以包含多个条件格式。每个格式由一个 [**FormatCondition** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#jsObject_FormatCondition)对象代表。

有关条件格式的详细信息，请参阅 [FormatCondition ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#jsObject_FormatCondition)对象。

使用 **FormatConditions** 属性可返回 **FormatConditions** 对象。使用 [**Add** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatConditions/FormatConditions%20.htm#FormatConditions.Add)方法可新建条件格式，使用 **Modify** 方法可更改现有的条件格式。

下例为 E1:E10 单元格添加条件格式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Add(xlCellValue, xlGreater, "=$a$1")     let range3 = range2.Borders     range3.LineStyle = xlContinuous     range3.Weight = xlThin     range3.ColorIndex = 6     let range4 = range2.Font     range4.Bold = true     range4.ColorIndex = 3 } ` |

**方法**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                 | 添加新的条件格式。返回一个 **FormatCondition**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#jsObject_FormatCondition)对象，它代表新的条件格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddAboveAverage**     | 返回一个代表指定区域条件格式规则的新 **AboveAverage** 对象。返回一个**AboveAverage** 对象 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddColorScale**       | 返回一个代表条件格式规则的新 **ColorScale** 对象，该条件格式规则使用单元格颜色渐变来表示选定区域中所含单元格值的相对差异。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddDatabar**          | 返回一个代表指定区域数据条条件格式规则的 **Databar** 对象。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddIconSetCondition** | 返回一个代表指定区域图标集条件格式规则的新 **IconSetCondition** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTop10**            | 返回一个代表指定区域条件格式规则的 **Top10** 对象。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddUniqueValues**     | 返回一个代表指定区域条件格式规则的新 **UniqueValues** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**              | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                | 从集合中返回一个对象。                                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **FormatConditions.Add**

添加新的条件格式。返回一个 **FormatCondition**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#jsObject_FormatCondition)对象，它代表新的条件格式。

**语法**

**express.Add(Type, Operator, Formula1, Formula2)**

*express*   一个代表 **FormatConditions** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**              | **说明**                                                     |
| ---------- | ------------- | ------------------------- | ------------------------------------------------------------ |
| *Type*     | 必选          | **XlFormatConditionType** | 指定条件格式是基于单元格值还是基于表达式。                   |
| *Operator* | 可选          | **Variant**               | 条件格式运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。如果 Type 为 xlExpression，则忽略 Operator 参数。 |
| *Formula1* | 可选          | **Variant**               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |
| *Formula2* | 可选          | **Variant**               | 当 Operator 为 xlBetween 或 xlNotBetween 时，它是与条件格式第二部分相关联的值或表达式（否则忽略该参数）。可为常量值、字符串值、单元格引用或公式。 |

**说明**

对单个区域定义的条件格式不能超过三个。使用 **Modify** 方法可修改现有的条件格式，使用 **Delete** 方法可在添加新条件格式前删除现有的格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例向单元格区域 E1:E10 中添加条件格式*/ function test() {     let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Add(xlCellValue, xlGreater, "=$a$1")     let range3 = range2.Borders     range3.LineStyle = xlContinuous     range3.Weight = xlThin     range3.ColorIndex = 6     let range4 = range2.Font     range4.Bold = true     range4.ColorIndex = 3 }` |

#### **FormatConditions.AddAboveAverage**

返回一个代表指定区域条件格式规则的新 **AboveAverage** 对象。返回一个**AboveAverage** 对象

**语法**

**express.AddAboveAverage()**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

**AboveAverage** 对象用于查找单元格区域中高于或低于平均或标准偏差的值。例如，可在年度业绩评估中查找高于平均业绩的人员，或者在质量评估中查找低于两个标准偏差的加工原料。

#### **FormatConditions.AddColorScale**

返回一个代表条件格式规则的新 **ColorScale** 对象，该条件格式规则使用单元格颜色渐变来表示选定区域中所含单元格值的相对差异。

**语法**

**express.AddColorScale(ColorScaleType)**

*express*   一个代表 **FormatConditions** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**     |
| ---------------- | ------------- | ------------ | ------------ |
| *ColorScaleType* | 必选          | **Long**     | 色阶的类型。 |

**说明**

色阶是直观的参照，可以帮助您了解数据的分布和变化。色阶可帮助您利用颜色的变化来标出给定区域中单元格值的相对差异。不同的颜色和颜色之间的渐变代表单元格值的差异。例如，在三色色阶中，您可以指定包含最高相对数据值的单元格为绿色，包含中间值的单元格为黄色，包含最低值的单元格为红色。

#### **FormatConditions.AddDatabar**

返回一个代表指定区域数据条条件格式规则的 **Databar** 对象。

**语法**

**express.AddDatabar()**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

数据条可帮助您查看某个单元格相对于其他单元格的值。数据条的长度代表单元格中的值。较长的数据条代表较高的值，较短的数据条代表较低的值。数据条在辨认较高和较低数值，特别是有大量数据存在的情况下很有帮助，比如，辨认假日销售报告中的最高玩具销量和最低玩具销量。

#### **FormatConditions.AddIconSetCondition**

返回一个代表指定区域图标集条件格式规则的新 **IconSetCondition** 对象。

**语法**

**express.AddIconSetCondition()**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

使用图标集为数据添加注释并将数据分为按阈值隔开的三到五类数据。每种图标均代表某一值范围。例如，在 3 箭头图标集中，红色向上箭头代表较高值，黄色侧箭头代表中间值，绿色向下箭头代表较低值。

#### **FormatConditions.AddTop10**

返回一个代表指定区域条件格式规则的 **Top10** 对象。

**语法**

**express.AddTop10()**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

使用 **Top10** 对象，可以基于指定的截止值查找单元格区域中的最高值和最低值。例如，可以查找区域报告中位居前五位的销售产品、客户调查中位居最后百分之十五的产品，或者部门人员分析中位居前 25 位的薪金。

#### **FormatConditions.AddUniqueValues**

返回一个代表指定区域条件格式规则的新 **UniqueValues** 对象。

**语法**

**express.AddUniqueValues()**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

可以使用 **UniqueValues** 对象快速显现包含唯一值或重复值的单元格。

#### **FormatConditions.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **FormatConditions** 对象的变量。

#### **FormatConditions.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **FormatConditions** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例为 E1:E10 单元格区域的现有条件格式设置格式属性*/ function test() {     let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)     let range3 = range2.Borders     range3.LineStyle = xlContinuous     range3.Weight = xlThin     range3.ColorIndex = 6 }` |

**成员属性**

#### **FormatConditions.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **FormatConditions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") }` |

#### **FormatConditions.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **FormatConditions** 对象的变量。

#### **FormatConditions.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FormatConditions** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **FormatConditions.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FormatConditions** 对象的变量。

适用环境：web

适用平台：windows/linux