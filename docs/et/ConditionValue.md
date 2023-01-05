**ConditionValue**



代表数据条条件格式规则计算最短数据条和最长数据条的方法。

**说明**

**ConditionValue** 对象是使用 **Databar** 对象的 **MaxPoint** 或 **MinPoint** 属性返回的。

通过使用 **Modify** 方法，您可以从默认设置（最低值表示最短的数据条，最高值表示最长的数据条）中更改计算类型。

以下示例将创建一个数据范围，然后对该范围应用数据条。您将会发现，由于范围中存在非常低和非常高的值，因此中间值具有长度相似的数据条。为了明确显示中间值，示例代码使用 **ConditionValue** 对象将阈值的计算方式更改为百分点。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){  //Create a range of data with a couple of extreme values     Application.ActiveSheet.Range("D1").Value2 = 1     Application.ActiveSheet.Range("D2").Value2 = 45     Application.ActiveSheet.Range("D3").Value2 = 50     Application.ActiveSheet.Range("D2:D3").AutoFill(Range("D2:D8"))     Application.ActiveSheet.Range("D9").Value2 = 500          Range("D1:D9").Select()              //Create a data bar with default behavior     let cfDataBar = Application.Selection.FormatConditions.AddDatabar()     alert("Because of the extreme values, middle data bars are very similar")          //The MinPoint and MaxPoint properties return a ConditionValue object     //which you can use to change threshold parameters     cfDataBar.MinPoint.Modify(xlConditionValuePercentile, 5)     cfDataBar.MaxPoint.Modify(xlConditionValuePercentile, 75)  }` |

**方法**

|                                                              | 名称       | 说明                                                     |
| ------------------------------------------------------------ | ---------- | -------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Modify** | 修改数据条条件格式规则计算最长数据条或最短数据条的方法。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回 **XlConditionValueTypes** 枚举的常量之一，该常量指定如何确定数据栏、色标或图标集条件格式的阈值。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 针对数据条条件格式返回或设置最短或最长的数据条的阈值。可读/写 **Variant** 类型。 |

**成员方法**

#### **ConditionValue.Modify**

修改数据条条件格式规则计算最长数据条或最短数据条的方法。

**语法**

**express.Modify(newtype, newvalue)**

*express*   一个代表 **ConditionValue** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**              | **说明**                                                     |
| ---------- | ------------- | ------------------------- | ------------------------------------------------------------ |
| *newtype*  | 必选          | **XlConditionValueTypes** | 指定最短数据条或最长数据条的计算方法。最短数据条的默认值是 xlConditionLowestValue；最长数据条的默认值是 xlConditionHighestValue。 |
| *newvalue* | 可选          | **Variant**               | 分配给最短或最长数据条的值。取决于 newtype 参数，此值可以是数字，也可以是计算结果为数字的公式。 |

**说明**

table { word-break:break-all; }

下表描述了每种计算类型的可接受阈值。

| newtype 参数               | newvalue 参数           |
| -------------------------- | ----------------------- |
| xlConditionLowestValue     | 忽略参数                |
| xlConditionHighestValue    | 忽略参数                |
| xlConditionValueNumber     | 任何数字                |
| xlConditionValuePercent    | 0 到 100 之间的任何数字 |
| xlConditionValuePercentile | 0 到 100 之间的任何数字 |
| xlConditionValueFormula    | 一个返回单一数字的公式  |

**成员属性**

#### **ConditionValue.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ConditionValue** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **ConditionValue.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ConditionValue** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ConditionValue.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ConditionValue** 对象的变量。

#### **ConditionValue.Type**

返回 **XlConditionValueTypes** 枚举的常量之一，该常量指定如何确定数据栏、色标或图标集条件格式的阈值。只读。

**语法**

**express.Type**

*express*   一个代表 **ConditionValue** 对象的变量。

#### **ConditionValue.Value**

针对数据条条件格式返回或设置最短或最长的数据条的阈值。可读/写 **Variant** 类型。

**语法**

**express.Value**

*express*   一个代表 **ConditionValue** 对象的变量。

**说明**

只有将条件格式的 **ConditionValue.Type** 属性设置为以下常量之一，您才可以设置值：**xlConditionValueNumber**、**xlConditionValuePercent**、**xlConditionValuePercentile** 或 **xlConditionValueFormula**。

如果阈值类型是公式，您可以将公式设置为 **String**。此公式必须返回单一数字。

适用环境：web

适用平台：windows/linux