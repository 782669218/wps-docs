**ErrorCheckingOptions**



代表应用程序的错误检查选项。

**说明**

使用 **Application** 对象的 **ErrorCheckingOptions** 属性可返回 **ErrorCheckingOptions** 对象。

引用 **Errors** 对象的 **Item** 属性可查看与错误检查选项关联的索引值列表。

返回 **ErrorCheckingOptions** 对象后，可使用以下属性设置或返回错误检查选项。这些属性都是 **ErrorCheckingOptions** 对象的成员。

- **BackgroundChecking**
- **EmptyCellReferences**
- **EvaluateToError**
- **InconsistentFormula**
- **IndicatorColorIndex**
- **NumberAsText**
- **OmittedCells**
- **TextDate**
- **UnlockedFormulaCells**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例使用 TextDate 属性启用对两位数年份文本日期的错误检查，并通知用户。*/ function test(){  let rngFormula = Application.Range("A1")      Range("A1").Formula = "'April 23, 00"     Application.ErrorCheckingOptions.TextDate = true      //Perform check to see if 2 digit year TextDate check is on.     if(rngFormula.Errors.Item(xlTextDate).Value == true) {         alert("The text date error checking feature is enabled.")     }     else {         alert("The text date error checking feature is not on.")     }  }` |

**属性**

|                                                              | 名称                         | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**              | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundChecking**       | 警告用户并指出违背可用错误检查规则的所有单元格。如果该属性设置为 **True** （默认值），则在违背启用错误的所有单元格旁边将出现**“自动更正选项”**按钮。如果为 **False**，则禁用背景错误检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                  | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmptyCellReferences**      | 如果该属性设置为 **True**（默认值），则 ET 使用**“自动更正选项”**按钮识别被选取的单元格，这些单元格中包括引用了空单元格的公式。如果该值为 **False**，则禁用空单元格引用检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EvaluateToError**          | 如果该属性设置为 **True**（默认值），则 ET 将使用**“自动更正选项”**按钮识别被选中的单元格，这些单元格中包含计算结果错误的公式。如果该值为 **False**，则禁用对计算结果为错误值的单元格的错误检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InconsistentFormula**      | 如果该值为 **True**（默认值），则 ET 将识别包含不一致公式的单元格区域。如果该值为 **False**，则禁用不一致公式的检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InconsistentTableFormula** | 如果表公式不一致，则返回 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IndicatorColorIndex**      | 返回或设置错误检查选项的指示器的颜色。**XlColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListDataValidation**       | 如果在列表中启用了数据有效性验证，则该属性值为 **Boolean** 值 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberAsText**             | 如果该值为 **True**（默认值），则 ET 将用**“自动更正选项”**按钮识别被选定的单元格，这些单元格包含文本格式的数字。如果该值为 **False**，则禁用对文本格式数字的错误检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OmittedCells**             | 如果该值为 **True**（默认值），则 ET 将用“自动更正选项”按钮识别包含公式的选定单元格，其中该公式引用的区域应包括相邻单元格，但这些单元格却被遗漏掉了。如果该值为 **False**，则禁用对被遗漏单元格的错误检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                   | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextDate**                 | 如果设置为 **True**（默认值），则 ET 将使用**“自动更正选项”**按钮识别包含用两位数字表示年份的文本日期的单元格。如果该属性值为 **False**，则禁用此类单元格的错误检查。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UnlockedFormulaCells**     | 如果该值为 **True**（默认值），则 ET 将识别未锁定并包含一个公式的选定单元格。如果该值为 **False**，则禁用包含公式的未锁定单元格的错误检查。**Boolean** 类型，可读写。 |

**成员属性**

#### **ErrorCheckingOptions.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }	} }` |

#### **ErrorCheckingOptions.BackgroundChecking**

警告用户并指出违背可用错误检查规则的所有单元格。如果该属性设置为 **True** （默认值），则在违背启用错误的所有单元格旁边将出现**“自动更正选项”**按钮。如果为 **False**，则禁用背景错误检查。**Boolean** 类型，可读写。

**语法**

**express.BackgroundChecking**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**说明**

请参考 **ErrorCheckingOptions** 对象以查阅该对象的可用成员列表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在下例中，当用户选择单元格 A1（该单元格包含引用空单元格的公式）时，将出现“自动更正选项”按钮。*/ function test(){  //Simulate an error by referring to empty cells.     Application.ErrorCheckingOptions.BackgroundChecking = true     Range("A1").Select()     ActiveCell.Formula = "=A2/A3" }` |

#### **ErrorCheckingOptions.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ErrorCheckingOptions.EmptyCellReferences**

如果该属性设置为 **True**（默认值），则 ET 使用**“自动更正选项”**按钮识别被选取的单元格，这些单元格中包括引用了空单元格的公式。如果该值为 **False**，则禁用空单元格引用检查。**Boolean** 类型，可读写。

**语法**

**express.EmptyCellReferences**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在下例中，单元格 A1 旁出现“自动更正选项”按钮，该单元格包括引用空单元格的公式。*/ function test(){  Application.ErrorCheckingOptions.EmptyCellReferences = true     Range("A1").Formula = "=A2+A3" }` |

#### **ErrorCheckingOptions.EvaluateToError**

如果该属性设置为 **True**（默认值），则 ET 将使用**“自动更正选项”**按钮识别被选中的单元格，这些单元格中包含计算结果错误的公式。如果该值为 **False**，则禁用对计算结果为错误值的单元格的错误检查。**Boolean** 类型，可读写。

**语法**

**express.EvaluateToError**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在本示例中，单元格 A3 旁出现“自动更正选项”按钮，该单元格包括一个被零除的错误*/ function test(){    //Simulate a divide-by-zero error.     Application.ErrorCheckingOptions.EvaluateToError = true     Range("A1").Value2 = 1     Range("A2").Value2 = 0     Range("A3").Formula = "=A1/A2" }` |

#### **ErrorCheckingOptions.InconsistentFormula**

如果该值为 **True**（默认值），则 ET 将识别包含不一致公式的单元格区域。如果该值为 **False**，则禁用不一致公式的检查。**Boolean** 类型，可读写。

**语法**

**express.InconsistentFormula**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**说明**

单元格区域中相一致的公式必须位于包含不一致的公式的单元格的左边、右边或上边、下边，以便 **InconsistentFormula** 属性正常工作。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在下例中，当用户选择单元格 B4（包含不一致的公式）时，将显示“自动更正选项”按钮。*/ function test(){   Application.ErrorCheckingOptions.InconsistentFormula = true      Range("A1:A3").Value2 = 1     Range("B1:B3").Value2 = 2     Range("C1:C3").Value2 = 3      Range("A4").Formula = "=SUM(A1:A3)"  //Consistent formula.     Range("B4").Formula = "=SUM(B1:B2)"  //Inconsistent formula.     Range("C4").Formula = "=SUM(C1:C3)"  //Consistent formula. }` |

#### **ErrorCheckingOptions.InconsistentTableFormula**

如果表公式不一致，则返回 **True**。可读/写 **Boolean** 类型。

**语法**

**express.InconsistentTableFormula**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

#### **ErrorCheckingOptions.IndicatorColorIndex**

返回或设置错误检查选项的指示器的颜色。**XlColorIndex** 类型，可读写。

**语法**

**express.IndicatorColorIndex**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**说明**

| **XlColorIndex** 可以是下列 **XlColorIndex** 常量之一。 |
| ------------------------------------------------------- |
| **xlColorIndexAutomatic** 为默认系统颜色。              |
| **xlColorIndexNone** 不与该属性一起使用。               |

通过输入相应的索引值可为指示器指定特殊的颜色。可用 **Colors** 属性返回当前的调色板。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在本示例中，ET 将查看错误检查的指示器颜色是否设置为默认系统颜色，并通知给用户。*/ function test(){  if(Application.ErrorCheckingOptions.IndicatorColorIndex == xlColorIndexAutomatic) {         alert("Your indicator color for error checking is set to the default system color.")     }     else {         alert("Your indicator color for error checking is not set to the default system color.")     } }` |

#### **ErrorCheckingOptions.ListDataValidation**

如果在列表中启用了数据有效性验证，则该属性值为 **Boolean** 值 **True**。**Boolean** 类型，可读写。

**语法**

**express.ListDataValidation**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

#### **ErrorCheckingOptions.NumberAsText**

如果该值为 **True**（默认值），则 ET 将用**“自动更正选项”**按钮识别被选定的单元格，这些单元格包含文本格式的数字。如果该值为 **False**，则禁用对文本格式数字的错误检查。**Boolean** 类型，可读写。

**语法**

**express.NumberAsText**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*在下例中，包含文本格式数字的单元格 A1 旁出现“自动更正选项”按钮。*/ function test(){  //Simulate an error by referencing a number stored as text.     Application.ErrorCheckingOptions.NumberAsText = true     Range("A1").Value = "'1" }` |

#### **ErrorCheckingOptions.OmittedCells**

如果该值为 **True**（默认值），则 ET 将用“自动更正选项”按钮识别包含公式的选定单元格，其中该公式引用的区域应包括相邻单元格，但这些单元格却被遗漏掉了。如果该值为 **False**，则禁用对被遗漏单元格的错误检查。**Boolean** 类型，可读写。

**语法**

**express.OmittedCells**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例中，“自动更正选项”按钮出现在包含公式的单元格 A4 旁。*/ function test(){  Application.ErrorCheckingOptions.OmittedCells = true     Range("A1").Value2 = 1     Range("A2").Value2 = 2     Range("A3").Value2 = 3     Range("A4").Formula = "=Sum(A1:A2)" }` |

#### **ErrorCheckingOptions.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

#### **ErrorCheckingOptions.TextDate**

如果设置为 **True**（默认值），则 ET 将使用**“自动更正选项”**按钮识别包含用两位数字表示年份的文本日期的单元格。如果该属性值为 **False**，则禁用此类单元格的错误检查。**Boolean** 类型，可读写。

**语法**

**express.TextDate**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例中，“自动更正选项”按钮将出现在单元格 A1 旁，该单元格包含用两位数字表示年份的文本日期。*/  function test(){   //Simulate an error by referencing a text date with a two-digit year.     Application.ErrorCheckingOptions.TextDate = true     Range("A1").Formula = "'April 23, 00" }` |

#### **ErrorCheckingOptions.UnlockedFormulaCells**

如果该值为 **True**（默认值），则 ET 将识别未锁定并包含一个公式的选定单元格。如果该值为 **False**，则禁用包含公式的未锁定单元格的错误检查。**Boolean** 类型，可读写。

**语法**

**express.UnlockedFormulaCells**

*express*   一个代表 **ErrorCheckingOptions** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例中，“自动更正选项”按钮出现在单元格 A3 旁，该单元格是一个包含公式的未锁定单元格。*/ function test(){  Application.ErrorCheckingOptions.UnlockedFormulaCells = true     Range("A1").Value2 = 1     Range("A2").Value2 = 2     Range("A3").Formula = "=A1+A2"     Range("A3").Locked = false }` |

适用环境：web

适用平台：windows/linux