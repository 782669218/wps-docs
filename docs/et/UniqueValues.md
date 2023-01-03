**UniqueValues**



**UniqueValues** 对象使用 **DupeUnique** 属性返回或设置一个枚举，该枚举确定规则是查找区域中的重复值还是唯一值。

**方法**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**               | 删除指定的条件格式规则对象。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ModifyAppliesToRange** | 设置此格式规则所应用于的单元格区域。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetLastPriority**      | 为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AppliesTo**    | 返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**      | 返回一个 **Borders** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DupeUnique**   | 返回或设置 **XlDupeUnique** 枚举的常量之一，指定条件格式规则是查找唯一的值还是查找重复的值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**         | 返回一个 **Font** 对象，该对象在条件格式规则的计算结果为 **True** 时指定字体格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Interior**     | 返回一个 **Interior** 对象，该对象在条件格式规则的计算结果为 **True** 时指定单元格的内部属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat** | 在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**     | 返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PTCondition**  | 返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表图表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScopeType**    | 返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StopIfTrue**   | 返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**         | 返回 **XlFormatConditionType** 枚举的常量之一，该常量指定条件格式的类型。只读。 |

**成员方法**

#### **UniqueValues.Delete**

删除指定的条件格式规则对象。

**语法**

**express.Delete()**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.ModifyAppliesToRange**

设置此格式规则所应用于的单元格区域。

**语法**

**express.ModifyAppliesToRange(Range)**

*express*   一个代表 **UniqueValues** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Range*  | 必选          | **Range**    | 此格式规则将应用于的区域。 |

**说明**

该区域必须采用 A1 引用样式，并全部包含在 **FormatConditions** 集合的父工作表内。可包括区域操作符（冒号）、相交区域操作符（空格）或合并区域操作符（逗号）。也可以使用货币符号，但会被忽略。

您也可以在区域的任意部分中使用局部定义名称，但该名称必须使用宏语言。

#### **UniqueValues.SetLastPriority**

为此条件格式规则设置求值顺序，以便在工作表上的所有其他规则之后计算此规则。

**语法**

**express.SetLastPriority()**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

优先级的实际值将等于工作表上条件格式规则的总数。当工作表中有多个条件格式规则时，此方法将导致优先级值大于此规则的规则的优先级增加 1。

**注释：**

**成员属性**

#### **UniqueValues.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **UniqueValues.AppliesTo**

返回一个 **Range** 对象，指定格式规则将应用于的单元格区域。

**语法**

**express.AppliesTo**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.Borders**

返回一个 **Borders** 集合，该集合在条件格式规则的计算结果为 **True** 时指定单元格边框的格式。只读。

**语法**

**express.Borders**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

对于条件格式对象，您只能为单元格的上边框、下边框和侧边框设置属性。

#### **UniqueValues.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **UniqueValues.DupeUnique**

返回或设置 **XlDupeUnique** 枚举的常量之一，指定条件格式规则是查找唯一的值还是查找重复的值。

**语法**

**express.DupeUnique**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.Font**

返回一个 **Font** 对象，该对象在条件格式规则的计算结果为 **True** 时指定字体格式。只读。

**语法**

**express.Font**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

条件格式对象并不支持 **Font** 对象的所有属性。您可以设置字体样式、下划线、颜色和删除线属性。

#### **UniqueValues.Interior**

返回一个 **Interior** 对象，该对象在条件格式规则的计算结果为 **True** 时指定单元格的内部属性。只读。

**语法**

**express.Interior**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.NumberFormat**

在条件格式规则的计算结果为 **True** 时返回或设置应用于单元格的数字格式。**Variant** 型，可读写。

**语法**

**express.NumberFormat**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

数字格式是使用**“单元格格式”**对话框的**“数字”**选项卡上显示的相同格式代码指定的。您可以使用内置的数字格式，例如 `"General"` 或者创建自定义数字格式。

#### **UniqueValues.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.Priority**

返回或设置条件格式规则的优先级值。当工作表中存在多个条件格式规则时，优先级确定求值的顺序。

**语法**

**express.Priority**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

设置优先级时，值必须为介于 1 和工作表上条件格式规则总数之间的正整数。对于工作表上的所有规则，优先级必须为唯一值，因此，如果更改指定条件格式规则的优先级，将可能会导致工作表上其他规则的优先级值移位。

#### **UniqueValues.PTCondition**

返回一个 **Boolean** 值，指明是否将条件格式应用于数据透视表图表。只读。

**语法**

**express.PTCondition**

*express*   一个代表 **UniqueValues** 对象的变量。

#### **UniqueValues.ScopeType**

返回或设置 **XlPivotConditionScope** 枚举的常量之一，该常量确定条件格式在应用于数据透视表图表时的范围。

**语法**

**express.ScopeType**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

默认值为 **xlSelectionScope**，它使用 **AppliesTo** 属性设置范围。

#### **UniqueValues.StopIfTrue**

返回或设置一个 **Boolean** 值，该值确定在当前规则的计算结果为 **True** 时是否应计算单元格上的其他格式规则。

**语法**

**express.StopIfTrue**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

为了支持向后兼容性，此属性的默认值为 **True**，而这与在用户界面中创建的格式规则的行为相反。

#### **UniqueValues.Type**

返回 **XlFormatConditionType** 枚举的常量之一，该常量指定条件格式的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **UniqueValues** 对象的变量。

**说明**

此属性将始终返回 **Long** 值“8”，等价于 **xlUniqueValues**。

适用环境：web

适用平台：windows/linux