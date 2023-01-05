**FormatCondition**



代表条件格式。

**说明**

**FormatCondition** 对象是 [**FormatConditions** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatConditions/FormatConditions%20.htm#jsObject_FormatConditions)集合的成员。对于给定区域，**FormatConditions** 集合中包含的条件格式不能超过三个。

使用 **Add** 方法可新建条件格式。如果区域内存在多种格式，则可使用 **Modify** 方法更改其中一种格式，或使用 **Delete** 方法删除一种格式，然后使用 **Add** 方法创建一种新格式。

使用 **FormatCondition** 对象的 **Font**、[**Borders** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#FormatCondition.Borders)和 [**Interior** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatCondition/FormatCondition%20.htm#FormatCondition.Interior)属性可控制已设置格式的单元格的外观。条件格式对象模型不支持这些对象的某些属性。下表列出所有可与条件格式一起使用的属性。

| 对象         | 属性                                                         |
| ------------ | ------------------------------------------------------------ |
| **Font**     | **Bold****Color****ColorIndex****FontStyle****Italic****Strikethrough****Underline**无法使用会计用下划线样式。 |
| **Border**   | **Bottom****Color****Left****Right****Style**可使用下列边框样式（其他均不可用）：**xlNone**、**xlSolid**、**xlDash**、**xlDot**、**xlDashDot**、**xlDashDotDot**、**xlGray50**、**xlGray75** 和 **xlGray25**。**Top****Weight**可使用下列边框粗细（其他均不可用）：**xlWeightHairline** 和 **xlWeightThin**。 |
| **Interior** | **Color****ColorIndex****Pattern****PatternColorIndex**      |

使用 **FormatConditions**(*index*)（其中 *index* 为条件格式的索引号）可返回 **FormatCondition** 对象。下例为 E1:E10 单元格的现有条件格式设置格式属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)     let range3 = range2.Borders     range3.LineStyle = xlContinuous     range3.Weight = xlThin     range3.ColorIndex = 6     let range4 = range2.Font     range4.Bold = true     range4.ColorIndex = 3 }` |

**方法**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**               | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Modify**               | 更改现有条件格式。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ModifyAppliesToRange** | 设置此格式规则所应用于的单元格区域。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetFirstPriority**     | 将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetLastPriority**      | 为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AppliesTo**    | 返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**      | 返回一个 **Borders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateOperator** | 指定格式条件中所用的日期运算符。可读/写。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**         | 返回一个 **Font** 对象，它代表指定对象的字体。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula1**     | 返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula2**     | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Interior**     | 返回一个 **Interior** 对象，它代表指定对象的内部。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat** | 在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Operator**     | 返回一个 **Long** 值，它代表条件格式的操作符。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PTCondition**  | 返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**     | 返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScopeType**    | 返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StopIfTrue**   | 返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**         | 返回或设置一个 **String** 类型的值，该值指定条件格式规则所用的文本字符串。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextOperator** | 返回或设置一个 **XlContainsOperator** 枚举常量，该常量指定条件格式规则执行的文本搜索。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**         | 返回一个包含 **xlFormatConditionType** 值的 **Long** 值，它代表对象类型。 |

**成员方法**

#### **FormatCondition.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Modify**

更改现有条件格式。

**语法**

**express.Modify(Type, Operator, Formula1, Formula2)**

*express*   一个代表 **FormatCondition** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**              | **说明**                                                     |
| ---------- | ------------- | ------------------------- | ------------------------------------------------------------ |
| *Type*     | 必选          | **XlFormatConditionType** | 指定条件格式是基于单元格值还是基于表达式。                   |
| *Operator* | 可选          | **Variant**               | 一个 XlFormatConditionOperator 值，它代表条件格式运算符。如果 Type 设为 xlExpression，则忽略该参数。 |
| *Formula1* | 可选          | **Variant**               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |
| *Formula2* | 可选          | **Variant**               | 与条件格式相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例修改单元格区域 E1:E10 的现有条件格式*/ Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1).Modify(xlCellValue, xlLess, "=$a$1")` |

#### **FormatCondition.ModifyAppliesToRange**

设置此格式规则所应用于的单元格区域。

**语法**

**express.ModifyAppliesToRange(Range)**

*express*   一个代表 **FormatCondition** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Range*  | 必选          | **Range**    | 此格式规则将应用于的区域。 |

**说明**

该区域必须采用 A1 引用样式，并全部包含在 [**FormatConditions** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatConditions/FormatConditions%20.htm#jsObject_FormatConditions)集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

#### **FormatCondition.SetFirstPriority**

将此条件格式规则的优先级值设置为“1”，以便在工作表上的所有其他规则之前计算此规则。

**语法**

**express.SetFirstPriority()**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

当工作表中有多个条件格式规则时，如果之前未将此规则设置为优先级“1”，此方法将导致工作表上的所有其他现有规则的优先级都增加 1。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 将在工作表级别应用条件格式规则的优先级。             |

#### **FormatCondition.SetLastPriority**

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

**语法**

**express.SetLastPriority()**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级增加 1。

| ![img](https://qn.cache.wpscdn.cn/gif/close.gif)注释 |
| ---------------------------------------------------- |
| 将在工作表级别应用条件格式规则的优先级。             |

**成员属性**

#### **FormatCondition.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **FormatCondition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息 let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") }` |

#### **FormatCondition.AppliesTo**

返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。

**语法**

**express.AppliesTo**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Borders**

返回一个 **Borders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

**语法**

**express.Borders**

*express*   一个代表 **FormatCondition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框*/ function SetRangeBorder() {          let range2 = Application.Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)     range2.LineStyle = xlContinuous     range2.Weight = xlThin     range2.ColorIndex = 3  }` |

#### **FormatCondition.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **FormatCondition.DateOperator**

指定格式条件中所用的日期运算符。可读/写。

**语法**

**express.DateOperator**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Font**

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Formula1**

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String** 型，只读。

**语法**

**express.Formula1**

*express*   一个代表 **FormatCondition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果单元格区域 E1:E10 的第一个条件格式的公式指定的是“小于 5”，则此示例将对其进行更改*/ let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)     if(range2.Operator == xlLess && range2.Formula1 == "5") {         range2.Modify(xlCellValue, xlLess, "10")     }` |

#### **FormatCondition.Formula2**

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。

**语法**

**express.Formula2**

*express*   一个代表 **FormatCondition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果单元格区域 E1:E10 的第一个条件格式的公式指定的是“介于 5 和 10 之间”，则本示例将对其进行更改*/ let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)     if(range2.Operator == xlBetween && range2.Formula1 == "5" && range2.Formula2 == "10") {         range2.Modify(xlCellValue, xlLess, "10")     }` |

#### **FormatCondition.Interior**

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.NumberFormat**

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。**Variant** 型，可读写。

**语法**

**express.NumberFormat**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

数字格式是使用**“单元格格式”**对话框的**“数字”**选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 `"General"` 或者创建自定义数字格式。

#### **FormatCondition.Operator**

返回一个 **Long** 值，它代表条件格式的操作符。

**语法**

**express.Operator**

*express*   一个代表 **FormatCondition** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果单元格区域 E1:E10 的条件格式一的公式指定的是“小于 5”，则此示例将对其进行更改*/ let range2 = Application.Worksheets.Item(1).Range("e1:e10").FormatConditions.Item(1)     if(range2.Operator == xlLess && range2.Formula1 == "5") {         range2.Modify(xlCellValue, xlBetween, "5", "15")     }` |

#### **FormatCondition.PTCondition**

返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表。只读。

**语法**

**express.PTCondition**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Priority**

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

**语法**

**express.Priority**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

#### **FormatCondition.ScopeType**

返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。

**语法**

**express.ScopeType**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

默认值为 **xlSelectionScope**，它使用 **AppliesTo** 属性设置范围。

#### **FormatCondition.StopIfTrue**

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

**语法**

**express.StopIfTrue**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

为了支持向后兼容性，此属性的默认值为 **True**，而这与在用户界面中创建的格式规则的行为相反。

#### **FormatCondition.Text**

返回或设置一个 **String** 类型的值，该值指定条件格式规则所用的文本字符串。

**语法**

**express.Text**

*express*   一个代表 **FormatCondition** 对象的变量。

**说明**

如果 **Type** 属性不是设置为 **xlTextString**，则忽略此属性。

#### **FormatCondition.TextOperator**

返回或设置一个 **XlContainsOperator** 枚举常量，该常量指定条件格式规则执行的文本搜索。

**语法**

**express.TextOperator**

*express*   一个代表 **FormatCondition** 对象的变量。

#### **FormatCondition.Type**

返回一个包含 **xlFormatConditionType** 值的 **Long** 值，它代表对象类型。

**语法**

**express.Type**

*express*   一个代表 **FormatCondition** 对象的变量。

适用环境：web

适用平台：windows/linux