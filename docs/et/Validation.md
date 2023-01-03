**Validation**



代表工作表区域的数据有效性规则。

**说明**

使用 **Validation** 属性可返回 **Validation** 对象。下例更改单元格 E5 的数据有效性规则。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Application.Range("e5").Validation.Modify(xlValidateList, xlValidAlertStop, null, "=$A$1:$A$10")` |

使用 **Add** 方法可将数据有效性添加到某个区域并创建一个新的 **Validation** 对象。此示例向单元格 E5 添加数据有效性检验。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let rng = Application.Range("e5").Validation rng.Add(xlValidateWholeNumber, xlValidAlertInformation, "5", "10") rng.InputTitle = "Integers" rng.ErrorTitle = "Integers" rng.InputMessage = "Enter an integer from five to ten" rng.ErrorMessage = "You must enter a number from five to ten"` |

**方法**

|                                                              | 名称       | 说明                             |
| ------------------------------------------------------------ | ---------- | -------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 向指定区域内添加数据有效性验证。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Modify** | 修改指定区域的数据有效性验证。   |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlertStyle**     | 返回有效性检验警告样式。**XlDVAlertStyle** 类型，只读。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorMessage**   | 返回或设置数据有效性检验错误消息。**String** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorTitle**     | 返回或设置数据有效性错误对话框的标题。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula1**       | 返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String** 型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula2**       | 返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IgnoreBlank**    | 如果指定区域内的数据有效性检验允许空值，则该值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IMEMode**        | 返回或设置日文输入规则的说明。可以是下表列出的 **XlIMEMode** 常量之一。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InCellDropdown** | 如果数据有效性显示含有有效取值的下拉列表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InputMessage**   | 返回或设置数据有效性检验输入信息。**String** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InputTitle**     | 返回或设置数据有效性输入对话框的标题。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Operator**       | 返回一个 **Long** 值，它代表数据有效性的运算符。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowError**      | 如果用户输入无效数据时显示数据有效性检查错误消息，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowInput**      | 如果用户在数据有效性检查区域内选定了某一单元格时，显示数据有效性检查输入消息，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**           | 返回一个包含 **XlDVType** 常量的 **Long** 值，它代表区域的数据类型有效性验证。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**          | 返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。 |

**成员方法**

#### **Validation.Add**

向指定区域内添加数据有效性验证。

**语法**

**express.Add(Type, AlertStyle, Operator, Formula1, Formula2)**

*express*   一个代表 **Validation** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*       | 必选          | **XlDVType** | 有效性验证类型。                                             |
| *AlertStyle* | 可选          | **Variant**  | 有效性验证警告的样式。可为以下 XlDVAlertStyle 常量之一：xlValidAlertInformation、xlValidAlertStop 或 xlValidAlertWarning。 |
| *Operator*   | 可选          | **Variant**  | 数据有效性验证运算符。可为以下 XlFormatConditionOperator 常量之一：xlBetween、xlEqual、xlGreater、xlGreaterEqual、xlLess、xlLessEqual、xlNotBetween 或 xlNotEqual。 |
| *Formula1*   | 可选          | **Variant**  | 数据有效性验证等式中的第一部分。                             |
| *Formula2*   | 可选          | **Variant**  | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性验证等式的第二部分（其他情况下，此参数被忽略）。 |

**说明**

**Add** 方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型                                               | 参数                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| **xlValidateCustom**                                         | **Formula1** 必需，忽略 **Formula2**。**Formula1** 必须包含一个表达式，数据项有效时该表达式的值为 **True**，数据项无效时，该值为 **False**。 |
| **xlInputOnly**                                              | 使用 **AlertStyle**、**Formula1** 或 **Formula2**。          |
| **xlValidateList**                                           | **Formula1** 必需，忽略 **Formula2**。**Formula1** 必须包含以逗号分隔的值列表，或对该列表的工作表引用。 |
| **xlValidateWholeNumber**、**xlValidateDate**、**xlValidateDecimal**、**xlValidateTextLength** 或 **xlValidateTime** | 必须指定 **Formula1** 或 **Formula2** 之一，或两者均指定。   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性验证。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = "Integers"     rng.ErrorTitle = "Integers"     rng.InputMessage = "Enter an integer from five to ten"     rng.ErrorMessage = "You must enter a number from five to ten" }` |

#### **Validation.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.Modify**

修改指定区域的数据有效性验证。

**语法**

**express.Modify(Type, AlertStyle, Operator, Formula1, Formula2)**

*express*   一个代表 **Validation** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*       | 可选          | **Variant**  | 一个代表有效性验证类型的 XlDVType 值。                       |
| *AlertStyle* | 可选          | **Variant**  | 一个代表有效性验证警告样式的 XlDVAlertStyle 值。             |
| *Operator*   | 可选          | **Variant**  | 一个代表数据有效性验证运算符的 XlFormatConditionOperator 值。 |
| *Formula1*   | 可选          | **Variant**  | 数据有效性验证等式中的第一部分。                             |
| *Formula2*   | 可选          | **Variant**  | 当 Operator 为 xlBetween 或 xlNotBetween 时，数据有效性的第二部分；其他情况下，此参数被忽略。 |

**说明**

**Modify** 方法所要求的参数依有效性验证的类型而定，如下表所示。

| 有效性验证类型                                               | 参数                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| **xlInputOnly**                                              | 不使用 **AlertStyle**、**Formula1** 和 **Formula2**。        |
| **xlValidateCustom**                                         | **Formula1** 必需，忽略 **Formula2**。**Formula1** 必须包含一个表达式，数据项有效时该表达式的值为 **True**，数据项无效时，该值为 **False**。 |
| **xlValidateList**                                           | **Formula1** 必需，忽略 **Formula2**。**Formula1** 必须包含一个以逗号分隔的值列表，或对该列表的工作表引用。 |
| **xlValidateDate**、**xlValidateDecimal**、**xlValidateTextLength**、 **xlValidateTime** 或 **xlValidateWholeNumber** | 必须指定 **Formula1** 或 **Formula2**，或两者均指定。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例更改单元格 E5 的数据有效性验证。 Application.Range("e5").Validation.Modify(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")` |

**成员属性**

#### **Validation.AlertStyle**

返回有效性检验警告样式。**XlDVAlertStyle** 类型，只读。

**语法**

**express.AlertStyle**

*express*   一个代表 **Validation** 对象的变量。

**说明**

使用 **Add** 方法设置某一区域的警告样式。如果该区域已有数据有效性检验，则可用 **Modify** 方法更改警告样式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示单元格 E5 的警告样式。 alert(Range("e5").Validation.AlertStyle)` |

#### **Validation.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Value == "ET")      {     	alert("This is an ET Application object.")     }     else      {     	alert("This is not an ET Application object.")     } }` |

#### **Validation.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Validation** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Validation.ErrorMessage**

返回或设置数据有效性检验错误消息。**String** 类型，可读写。

**语法**

**express.ErrorMessage**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = "Integers"     rng.ErrorTitle = "Integers"     rng.InputMessage = "Enter an integer from five to ten"     rng.ErrorMessage = "You must enter a number from five to ten" }` |

#### **Validation.ErrorTitle**

返回或设置数据有效性错误对话框的标题。**String** 类型，可读写。

**语法**

**express.ErrorTitle**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = "Integers"     rng.ErrorTitle = "Integers"     rng.InputMessage = "Enter an integer from five to ten"     rng.ErrorMessage = "You must enter a number from five to ten" }` |

#### **Validation.Formula1**

返回与条件格式或者数据有效性相关联的值或表达式。可为常量值、字符串值、单元格引用或公式。**String** 型，只读。

**语法**

**express.Formula1**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.Formula2**

返回与条件格式或数据有效性验证第二部分相关联的值或表达式。仅用于数据有效性条件格式 **Operator** 属性为 **xlBetween** 或 **xlNotBetween** 的情况。可为常量值、字符串值、单元格引用或公式。**String** 类型，只读。

**语法**

**express.Formula2**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.IgnoreBlank**

如果指定区域内的数据有效性检验允许空值，则该值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IgnoreBlank**

*express*   一个代表 **Validation** 对象的变量。

**说明**

当 **IgnoreBlank** 属性为 **True** 时，如果单元格为空白，或者 **MinVal** 或 **MaxVal** 属性所引用的单元格为空白，仍认为该单元格数据有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使单元格 E5 的数据有效性检验为允许有空值。 Application.Range("e5").Validation.IgnoreBlank = true` |

#### **Validation.IMEMode**

返回或设置日文输入规则的说明。可以是下表列出的 **XlIMEMode** 常量之一。**Long** 类型，可读写。

**语法**

**express.IMEMode**

*express*   一个代表 **Validation** 对象的变量。

**说明**

| 常量                      | 说明             |
| ------------------------- | ---------------- |
| **xlIMEModeAlpha**        | 半角字母数字     |
| **xlIMEModeAlphaFull**    | 全角字母数字     |
| **xlIMEModeDisable**      | 禁用             |
| **xlIMEModeHiragana**     | 平假名           |
| **xlIMEModeKatakana**     | 片假名           |
| **xlIMEModeKatakanaHalf** | 片假名（半角）   |
| **xlIMEModeNoControl**    | 无控制           |
| **xlIMEModeOff**          | 关闭（英文模式） |
| **xlIMEModeOn**           | 开启             |

只有安装和选择了日文语言支持时，才可设置该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置单元格 E5 的数据输入规则。 function test() {     let rng = Application.Range("E5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = " seisuuchi seisuuchi seisuuchi"     rng.ErrorTitle = " seisuuchi seisuuchi seisuuchi"     rng.InputMessage = "5  kara kara 10  否 seisuu seisuu o nyuuryoku nyuuryoku shite shite kudasai kudasai kudasai kudasai ."     rng.ErrorMessage = " nyuuryoku nyuuryoku dekirunowa dekirunowa dekirunowa dekirunowa dekirunowa 5  kara kara 10  madeno madeno madeno ne desu desu ."     rng.IMEMode = xlIMEModeAlpha }` |

#### **Validation.InCellDropdown**

如果数据有效性显示含有有效取值的下拉列表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.InCellDropdown**

*express*   一个代表 **Validation** 对象的变量。

**说明**

如果数据有效性检验类型不是 **xlValidateList**，将忽略该属性。

可用 **Validation** 对象的 **Add** 方法或 **Modify** 方法的 *Minimum* 参数来指定包含有效数据的区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性检验。单元格区域 A1:A10 包含单元格的有效取值，且该单元格显示含有这些有效取值的下拉列表。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateList, xlValidAlertStop, xlBetween, "=$A$1:$A$10")     rng.InCellDropdown = true }` |

#### **Validation.InputMessage**

返回或设置数据有效性检验输入信息。**String** 类型，可读写。

**语法**

**express.InputMessage**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = "Integers"     rng.ErrorTitle = "Integers"     rng.InputMessage = "Enter an integer from five to ten"     rng.ErrorMessage = "You must enter a number from five to ten" }` |

#### **Validation.InputTitle**

返回或设置数据有效性输入对话框的标题。**String** 类型，可读写。

**语法**

**express.InputTitle**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向单元格 E5 添加数据有效性检验，并指定输入信息和错误消息。 function test() {     let rng = Application.Range("e5").Validation     rng.Add(xlValidateWholeNumber,  xlValidAlertStop, xlBetween, "5", "10")     rng.InputTitle = "Integers"     rng.ErrorTitle = "Integers"     rng.InputMessage = "Enter an integer from five to ten"     rng.ErrorMessage = "You must enter a number from five to ten" }` |

#### **Validation.Operator**

返回一个 **Long** 值，它代表数据有效性的运算符。

**语法**

**express.Operator**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.ShowError**

如果用户输入无效数据时显示数据有效性检查错误消息，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowError**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一张工作表上的单元格 A10 中添加数据有效性检查。输入的值必须处于 5 到 10 之间；如果用户输入了无效数据，将显示错误消息，但不显示输入消息。 function test() {     let rng = Application.Worksheets.Item(1).Range("A10").Validation     rng.Add(1, xlValidAlertStop, xlBetween, "5", "10")     rng.ErrorMessage = "value must be between 5 and 10"     rng.ShowInput = false     rng.ShowError = true }` |

#### **Validation.ShowInput**

如果用户在数据有效性检查区域内选定了某一单元格时，显示数据有效性检查输入消息，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowInput**

*express*   一个代表 **Validation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一张工作表上的单元格 A10 中添加数据有效性检查。输入的值必须处于 5 到 10 之间；如果用户输入了无效数据，将显示错误消息，但不显示输入消息。 function test() {     let rng = Application.Worksheets.Item(1).Range("A10").Validation     rng.Add(1, xlValidAlertStop, xlBetween, "5", "10")     rng.ErrorMessage = "value must be between 5 and 10"     rng.ShowInput = false     rng.ShowError = true }` |

#### **Validation.Type**

返回一个包含 **XlDVType** 常量的 **Long** 值，它代表区域的数据类型有效性验证。

**语法**

**express.Type**

*express*   一个代表 **Validation** 对象的变量。

#### **Validation.Value**

返回一个 **Boolean** 值，它指明是否符合所有的有效性验证条件（即，该区域是否包含有效数据）。

**语法**

**express.Value**

*express*   一个代表 **Validation** 对象的变量。

适用环境：web

适用平台：windows/linux